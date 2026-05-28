# SPDX-License-Identifier: MIT
"""
Zemax Lens to SolidWorks Solid Model
=====================================
Reads lens surfaces from a Zemax OpticStudio lens file via the ZOS-API,
computes the aspheric sag profile as dense point arrays, then uses the
SolidWorks COM API to sketch the half-profile and revolve it into a solid.

Requirements:
  - Windows (COM automation required)
  - pythonnet (clr) for ZOS-API
  - pywin32 (win32com) for SolidWorks COM automation
  - OpticStudio installed with valid API license
  - SolidWorks already running

Usage:
  python zemax_to_solidworks_public.py <lens_file.zmx> --front 2 --back 3
  python zemax_to_solidworks_public.py <lens_file.zmx> --front 7 --back 8 --points 200
  python zemax_to_solidworks_public.py <lens_file.zmx> --front 2 --back 3 --template "C:\\path\\to\\Part.prtdot"

Configuration:
  Place a 'zemax_to_solidworks.json' file next to this script (or in your
  home directory) to set defaults for template path, points, etc.

  Example zemax_to_solidworks.json:
  {
    "sw_template": "C:\\Path\\To\\Your\\Part(mm).prtdot",
    "num_points": 150,
    "default_front_surface": 2,
    "default_back_surface": 3
  }
"""

import argparse
import clr
import json
import math
import os
import winreg

import numpy as np
import pythoncom
import win32com.client
from win32com.client import VARIANT


# =============================================================================
# Configuration Loading
# =============================================================================
CONFIG_FILENAME = "zemax_to_solidworks.json"

DEFAULT_CONFIG = {
    "sw_template": None,          # None = auto-detect from SolidWorks
    "num_points": 150,
    "default_front_surface": 2,
    "default_back_surface": 3,
}


def find_config_file():
    """Search for config file next to script, then in home directory."""
    # Next to this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(script_dir, CONFIG_FILENAME)
    if os.path.isfile(candidate):
        return candidate

    # Home directory
    home = os.path.expanduser("~")
    candidate = os.path.join(home, CONFIG_FILENAME)
    if os.path.isfile(candidate):
        return candidate

    return None


def load_config():
    """Load configuration from JSON file, falling back to defaults."""
    config = dict(DEFAULT_CONFIG)
    config_path = find_config_file()
    if config_path:
        try:
            with open(config_path, "r") as f:
                user_config = json.load(f)
            config.update(user_config)
            print(f"  Config loaded from: {config_path}")
        except (json.JSONDecodeError, OSError) as e:
            print(f"  Warning: Could not load config from {config_path}: {e}")
    return config


# =============================================================================
# Zemax ZOS-API Connection Class
# =============================================================================
class PythonStandaloneApplication(object):
    """Connects to OpticStudio via the ZOS-API .NET interface."""

    class LicenseException(Exception):
        pass

    class ConnectionException(Exception):
        pass

    class InitializationException(Exception):
        pass

    class SystemNotPresentException(Exception):
        pass

    def __init__(self, path=None):
        # Locate Zemax installation via Windows registry
        aKey = winreg.OpenKey(
            winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER),
            r"Software\Zemax", 0, winreg.KEY_READ
        )
        zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
        NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
        winreg.CloseKey(aKey)
        clr.AddReference(NetHelper)
        import ZOSAPI_NetHelper

        if path is None:
            isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize()
        else:
            isInitialized = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(path)

        if isInitialized:
            dir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory()
        else:
            raise self.InitializationException("Unable to locate Zemax OpticStudio.")

        clr.AddReference(os.path.join(os.sep, dir, "ZOSAPI.dll"))
        clr.AddReference(os.path.join(os.sep, dir, "ZOSAPI_Interfaces.dll"))
        import ZOSAPI

        self.ZOSAPI = ZOSAPI
        self.TheConnection = ZOSAPI.ZOSAPI_Connection()
        if self.TheConnection is None:
            raise self.ConnectionException("Unable to initialize .NET connection to ZOSAPI")

        self.TheApplication = self.TheConnection.CreateNewApplication()
        if self.TheApplication is None:
            raise self.InitializationException("Unable to acquire ZOSAPI application")

        if not self.TheApplication.IsValidLicenseForAPI:
            raise self.LicenseException("License is not valid for ZOSAPI use")

        self.TheSystem = self.TheApplication.PrimarySystem
        if self.TheSystem is None:
            raise self.SystemNotPresentException("Unable to acquire Primary system")

    def __del__(self):
        if self.TheApplication is not None:
            self.TheApplication.CloseApplication()
            self.TheApplication = None
        self.TheConnection = None

    def OpenFile(self, filepath, saveIfNeeded):
        if self.TheSystem is None:
            raise self.SystemNotPresentException("Unable to acquire Primary system")
        self.TheSystem.LoadFile(filepath, saveIfNeeded)


# =============================================================================
# Aspheric Sag Computation
# =============================================================================
def even_asphere_sag(r, radius, conic, coeffs):
    """
    Compute sag z(r) for an even asphere surface.

    z(r) = c*r^2 / (1 + sqrt(1 - (1+k)*c^2*r^2)) + sum(alpha_i * r^(2i))

    Parameters
    ----------
    r : float or ndarray
        Radial distance from optical axis
    radius : float
        Radius of curvature (R). Use float('inf') or 0 curvature for flat.
    conic : float
        Conic constant (k)
    coeffs : list of float
        Even asphere coefficients [alpha_2, alpha_4, alpha_6, ...] where
        alpha_i multiplies r^(2i). In Zemax "Even Asphere", Par1 = coeff
        on r^2, Par2 = coeff on r^4, etc.

    Returns
    -------
    z : float or ndarray
        Sag value(s)
    """
    r = np.asarray(r, dtype=float)
    z = np.zeros_like(r)

    # Curvature contribution
    if radius != 0 and not np.isinf(radius):
        c = 1.0 / radius
        k = conic
        denom_arg = 1.0 - (1.0 + k) * c**2 * r**2
        # Clamp to avoid sqrt of negative (at edge of validity)
        denom_arg = np.maximum(denom_arg, 1e-15)
        z = (c * r**2) / (1.0 + np.sqrt(denom_arg))

    # Aspheric polynomial terms
    for i, alpha in enumerate(coeffs):
        if alpha != 0:
            power = 2 * (i + 1)  # r^2, r^4, r^6, ...
            z = z + alpha * r**power

    return z


# =============================================================================
# Phase 1: Extract Lens Parameters from Zemax
# =============================================================================
def extract_lens_params(lens_file, surf_front, surf_back):
    """Extract surface parameters from the Zemax lens file."""
    print("=" * 60)
    print("Phase 1: Extracting lens parameters from Zemax")
    print("=" * 60)

    if not os.path.isfile(lens_file):
        raise FileNotFoundError(f"Lens file not found: {lens_file}")

    zos = PythonStandaloneApplication()
    ZOSAPI = zos.ZOSAPI
    TheSystem = zos.TheSystem

    # Load lens file
    TheSystem.LoadFile(lens_file, False)
    print(f"  Loaded: {lens_file}")

    TheLDE = TheSystem.LDE
    surfaces = {}

    for surf_num in [surf_front, surf_back]:
        surf = TheLDE.GetSurfaceAt(surf_num)
        surf_data = {
            'surface_num': surf_num,
            'type_name': surf.TypeName,
            'radius': surf.Radius,
            'conic': surf.Conic,
            'thickness': surf.Thickness,
            'semi_diameter': surf.SemiDiameter,
            'material': surf.Material,
            'coefficients': []
        }

        # Read aspheric coefficients (Par1 through Par8 for Even Asphere)
        if 'asphere' in surf_data['type_name'].lower():
            par_columns = [
                ZOSAPI.Editors.LDE.SurfaceColumn.Par1,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par2,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par3,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par4,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par5,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par6,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par7,
                ZOSAPI.Editors.LDE.SurfaceColumn.Par8,
            ]
            for col in par_columns:
                cell = surf.GetSurfaceCell(col)
                if cell is not None:
                    try:
                        val = cell.DoubleValue
                        surf_data['coefficients'].append(val)
                    except Exception:
                        try:
                            val = float(cell.Value)
                            surf_data['coefficients'].append(val)
                        except Exception:
                            surf_data['coefficients'].append(0.0)
                else:
                    surf_data['coefficients'].append(0.0)
        else:
            # Standard or plano surface — no aspheric terms
            surf_data['coefficients'] = [0.0] * 8

        surfaces[surf_num] = surf_data

    # Print summary
    for num, sd in surfaces.items():
        print(f"\n  Surface {num} ({sd['type_name']}):")
        print(f"    Radius     = {sd['radius']:.6f} mm")
        print(f"    Conic      = {sd['conic']:.6e}")
        print(f"    Thickness  = {sd['thickness']:.6f} mm")
        print(f"    Semi-Diam  = {sd['semi_diameter']:.6f} mm")
        print(f"    Material   = {sd['material']}")
        non_zero = [(i + 1, c) for i, c in enumerate(sd['coefficients']) if c != 0]
        if non_zero:
            print(f"    Aspheric coefficients:")
            for idx, val in non_zero:
                print(f"      Par{idx} (r^{2*idx}) = {val:.10e}")
        else:
            print(f"    Aspheric coefficients: (all zero — spherical)")

    # Clean up
    del zos
    zos = None

    return surfaces


# =============================================================================
# Phase 2: Compute Profile Points
# =============================================================================
def compute_profile_points(surfaces, surf_front, surf_back, num_points):
    """
    Compute the closed half-profile of the lens for revolution.

    Coordinate system (matching SolidWorks sketch on Right Plane):
      - X axis = optical axis (horizontal in sketch)
      - Y axis = radial direction (vertical in sketch)

    Each surface is sampled only to its own clear semi-diameter. If one
    surface has a smaller clear aperture than the lens OD, a horizontal
    (radial) flat line extends it to the OD before the vertical edge line.

    Profile traversal:
      1. Front surface: r=0 → r=front_SD
      2. Front flat (if front_SD < OD): radial line at constant sag to OD
      3. Edge line: vertical from front OD to back OD
      4. Back flat (if back_SD < OD): radial line at constant sag from OD to back_SD
      5. Back surface: r=back_SD → r=0
    """
    print("\n" + "=" * 60)
    print("Phase 2: Computing profile points")
    print("=" * 60)

    front = surfaces[surf_front]
    back = surfaces[surf_back]

    front_sd = front['semi_diameter']
    back_sd = back['semi_diameter']
    # Lens OD is the larger of the two clear semi-diameters
    semi_diam = max(front_sd, back_sd)
    thickness = front['thickness']  # axial distance from front to back surface

    print(f"  Lens OD (semi): {semi_diam:.4f} mm")
    print(f"  Front clear SD: {front_sd:.4f} mm")
    print(f"  Back clear SD:  {back_sd:.4f} mm")
    print(f"  Center thickness: {thickness:.4f} mm")

    # Sample each surface to its own clear semi-diameter
    r_front = np.linspace(0, front_sd, num_points)
    r_back = np.linspace(0, back_sd, num_points)

    sag_front = even_asphere_sag(r_front, front['radius'], front['conic'], front['coefficients'])
    sag_back = even_asphere_sag(r_back, back['radius'], back['conic'], back['coefficients'])

    # Build profile points as (x, y) where x = optical axis, y = radial
    # Front surface vertex is at origin. Back surface vertex is at (thickness, 0).
    profile_points = []

    # Segment 1: Front surface from r=0 to r=front_SD
    for i in range(num_points):
        profile_points.append((sag_front[i], r_front[i]))

    # Segment 2 (front flat): if front_SD < OD, horizontal line at sag_front edge
    front_has_flat = front_sd < semi_diam - 1e-9
    if front_has_flat:
        profile_points.append((sag_front[-1], semi_diam))

    # Segment 3: Vertical edge line at OD
    back_has_flat = back_sd < semi_diam - 1e-9
    x_back_edge = thickness + sag_back[-1]
    profile_points.append((x_back_edge, semi_diam))

    # Segment 4 (back flat): if back_SD < OD, horizontal line from OD to back_SD
    if back_has_flat:
        profile_points.append((x_back_edge, back_sd))

    # Segment 5: Back surface from r=back_SD to r=0
    for i in range(num_points - 1, -1, -1):
        profile_points.append((thickness + sag_back[i], r_back[i]))

    # Closure along optical axis is implicit (endpoints are on-axis).

    print(f"  Front has flat annulus: {front_has_flat}")
    print(f"  Back has flat annulus:  {back_has_flat}")
    print(f"  Total profile points: {len(profile_points)}")
    print(f"  Front edge sag: {sag_front[-1]:.6f} mm")
    print(f"  Back edge sag: {sag_back[-1]:.6f} mm")

    # Return extra info for sketch drawing
    profile_info = {
        'front_has_flat': front_has_flat,
        'back_has_flat': back_has_flat,
        'front_sd': front_sd,
        'back_sd': back_sd,
    }
    return profile_points, semi_diam, thickness, profile_info


# =============================================================================
# Geometry Helpers
# =============================================================================
def is_spherical(surf_data):
    """Return True if surface is purely spherical (no conic, no aspheric terms)."""
    if surf_data['conic'] != 0:
        return False
    if any(c != 0 for c in surf_data['coefficients']):
        return False
    if surf_data['radius'] == 0 or np.isinf(surf_data['radius']):
        return False  # plano — handled separately as a line
    return True


# =============================================================================
# Phase 3: Build SolidWorks Model
# =============================================================================
def resolve_sw_template(swApp, template_path):
    """
    Resolve the SolidWorks part template to use.

    Priority:
      1. Explicit path from config/CLI (if it exists on disk)
      2. SolidWorks user preference for default part template directory
      3. SolidWorks GetDocumentTemplate for a new part
    """
    # Option 1: explicit path
    if template_path and os.path.isfile(template_path):
        return template_path

    if template_path:
        print(f"  Warning: Configured template not found: {template_path}")

    # Option 2: search SolidWorks template directory for an mm-based part template
    try:
        tmpl_dir = swApp.GetUserPreferenceStringValue(7)  # swDefaultTemplatePart folder
        if tmpl_dir and os.path.isdir(tmpl_dir):
            for f in os.listdir(tmpl_dir):
                if f.lower().endswith('.prtdot'):
                    # Prefer one with 'mm' in the name
                    if 'mm' in f.lower():
                        return os.path.join(tmpl_dir, f)
            # If none with 'mm', use the first .prtdot found
            for f in os.listdir(tmpl_dir):
                if f.lower().endswith('.prtdot'):
                    return os.path.join(tmpl_dir, f)
    except Exception:
        pass

    # Option 3: ask SolidWorks for the default part template
    try:
        default = swApp.GetUserPreferenceStringValue(21)  # swDefaultTemplatePart
        if default and os.path.isfile(default):
            return default
    except Exception:
        pass

    raise RuntimeError(
        "Could not find a SolidWorks part template. Please set 'sw_template' in "
        f"your {CONFIG_FILENAME} or pass --template on the command line."
    )


def build_solidworks_model(profile_points, semi_diam, thickness, num_points, config, surfaces, surf_front, surf_back, profile_info):
    """
    Create a revolved solid in SolidWorks from the lens profile.
    Attaches to the already-running SolidWorks session.
    """
    print("\n" + "=" * 60)
    print("Phase 3: Building SolidWorks model")
    print("=" * 60)

    # --- Connect to running SolidWorks ---
    swApp = win32com.client.GetActiveObject("SldWorks.Application")
    swApp.Visible = True
    print("  Connected to running SolidWorks instance.")

    # --- Constants ---
    swUnitSystem_MMGS = 2

    # --- Resolve template ---
    defaultTemplate = resolve_sw_template(swApp, config.get("sw_template"))
    print(f"  Using template: {defaultTemplate}")

    swModel = swApp.NewDocument(defaultTemplate, 0, 0, 0)
    print(f"  NewDocument returned: {type(swModel).__name__} = {swModel}")

    # Get the active document reference
    swModel = swApp.ActiveDoc
    if swModel is None:
        raise RuntimeError("Failed to create new part document in SolidWorks!")
    try:
        title = swModel.GetTitle
        print(f"  New part created: {title}")
    except Exception:
        print(f"  New part created.")

    # --- Set units to millimeters (MMGS) ---
    swModelDocExt = swModel.Extension
    try:
        swModelDocExt.SetUserPreferenceInteger(175, 0, swUnitSystem_MMGS)
    except Exception:
        try:
            swModelDocExt.SetUserPreferenceInteger(11, 0, 0)
        except Exception:
            pass  # Units may already be mm from template
    print("  Units set to MMGS (millimeters).")

    # --- Select the Right YZ Plane for sketching ---
    swModel.ClearSelection2(True)
    plane_names = ["Right YZ Plane", "Right Plane", "Right", "RIGHT"]
    boolstatus = False
    for pname in plane_names:
        boolstatus = swModel.Extension.SelectByID2(pname, "PLANE", 0, 0, 0, False, 0, pythoncom.Nothing, 0)
        if boolstatus:
            print(f"  Plane selected: '{pname}'")
            break
    if not boolstatus:
        # Fallback: find Right/YZ plane in feature tree
        for i in range(50):
            try:
                feat = swModel.FeatureByPositionReverse(i)
                if feat is None:
                    break
                if feat.GetTypeName2 == "RefPlane":
                    name = feat.Name
                    if "right" in name.lower() or "yz" in name.lower():
                        feat.Select2(False, 0)
                        boolstatus = True
                        print(f"  Plane selected by traversal: '{name}'")
                        break
            except Exception:
                break
    print(f"  Plane selected: {boolstatus}")

    # --- Open sketch ---
    swModel.SketchManager.InsertSketch(True)
    skActive = swModel.SketchManager.ActiveSketch
    print(f"  Sketch opened. Active sketch: {skActive}")

    # Enable direct database mode for batch sketch entity creation
    swModel.SketchManager.AddToDB = True

    # --- Draw the profile ---
    # Right YZ Plane sketch: sketch coords are (sketch_x, sketch_y, 0).
    # We map: radial → sketch_x (horizontal), optical_axis → sketch_y (vertical).
    # So profile point (optical_axis, radial) → sketch (radial, optical_axis, 0).
    # SolidWorks API uses meters.
    scale = 0.001  # mm to meters

    front_has_flat = profile_info['front_has_flat']
    back_has_flat = profile_info['back_has_flat']

    # Compute segment boundaries in the profile_points list
    # Segment 1: front surface = indices [0, num_points)
    # Segment 2: front flat (optional) = index num_points (if front_has_flat)
    # Segment 3: edge line target = next index
    # Segment 4: back flat (optional) = next index (if back_has_flat)
    # Segment 5: back surface = last num_points indices
    idx = 0
    front_pts = profile_points[idx:idx + num_points]
    idx += num_points

    if front_has_flat:
        front_flat_pt = profile_points[idx]
        idx += 1
    else:
        front_flat_pt = None

    edge_pt = profile_points[idx]  # far side of edge line at OD
    idx += 1

    if back_has_flat:
        back_flat_pt = profile_points[idx]
        idx += 1
    else:
        back_flat_pt = None

    back_pts = profile_points[idx:idx + num_points]

    x_extent = max(abs(profile_points[-1][0]), semi_diam) * 2.0
    front_data = surfaces[surf_front]
    back_data = surfaces[surf_back]

    # --- Segment 1: Front surface ---
    front_is_flat = all(abs(px - front_pts[0][0]) < 1e-12 for (px, py) in front_pts)

    if front_is_flat:
        seg = swModel.SketchManager.CreateLine(
            front_pts[0][1] * scale, front_pts[0][0] * scale, 0.0,
            front_pts[-1][1] * scale, front_pts[-1][0] * scale, 0.0
        )
        print(f"  Front surface line (plano): {seg}")
    elif is_spherical(front_data):
        R = front_data['radius']
        seg = swModel.SketchManager.CreateArc(
            0.0, R * scale, 0.0,                                    # center
            front_pts[0][1] * scale, front_pts[0][0] * scale, 0.0,  # start (vertex)
            front_pts[-1][1] * scale, front_pts[-1][0] * scale, 0.0,  # end (edge)
            1 if R > 0 else -1                                      # direction
        )
        print(f"  Front surface arc (R={R:.4f} mm): {seg}")
    else:
        spline_array_front = []
        for (px, py) in front_pts:
            spline_array_front.extend([py * scale, px * scale, 0.0])
        pt_data = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, spline_array_front)
        seg = swModel.SketchManager.CreateSpline2(pt_data, True)
        print(f"  Front surface spline ({num_points} pts): {seg}")

    # --- Segment 2: Front flat annulus (horizontal/radial line to OD) ---
    if front_has_flat:
        last_front = front_pts[-1]
        seg = swModel.SketchManager.CreateLine(
            last_front[1] * scale, last_front[0] * scale, 0.0,
            front_flat_pt[1] * scale, front_flat_pt[0] * scale, 0.0
        )
        print(f"  Front flat line (SD {profile_info['front_sd']:.3f} → OD {semi_diam:.3f}): {seg}")
        connect_front = front_flat_pt
    else:
        connect_front = front_pts[-1]

    # --- Segment 3: Vertical edge line at OD ---
    seg = swModel.SketchManager.CreateLine(
        connect_front[1] * scale, connect_front[0] * scale, 0.0,
        edge_pt[1] * scale, edge_pt[0] * scale, 0.0
    )
    print(f"  Edge line: {seg}")

    # --- Segment 4: Back flat annulus (horizontal/radial line from OD to back SD) ---
    if back_has_flat:
        seg = swModel.SketchManager.CreateLine(
            edge_pt[1] * scale, edge_pt[0] * scale, 0.0,
            back_flat_pt[1] * scale, back_flat_pt[0] * scale, 0.0
        )
        print(f"  Back flat line (OD {semi_diam:.3f} → SD {profile_info['back_sd']:.3f}): {seg}")
        connect_back = back_flat_pt
    else:
        connect_back = edge_pt

    # --- Segment 5: Back surface ---
    back_is_flat = all(abs(px - back_pts[0][0]) < 1e-12 for (px, py) in back_pts)

    if back_is_flat:
        seg = swModel.SketchManager.CreateLine(
            back_pts[0][1] * scale, back_pts[0][0] * scale, 0.0,
            back_pts[-1][1] * scale, back_pts[-1][0] * scale, 0.0
        )
        print(f"  Back surface line (plano): {seg}")
    elif is_spherical(back_data):
        R = back_data['radius']
        seg = swModel.SketchManager.CreateArc(
            0.0, (thickness + R) * scale, 0.0,                      # center
            back_pts[0][1] * scale, back_pts[0][0] * scale, 0.0,    # start (edge)
            back_pts[-1][1] * scale, back_pts[-1][0] * scale, 0.0,  # end (vertex)
            -1 if R > 0 else 1                                      # direction (reversed traverse)
        )
        print(f"  Back surface arc (R={R:.4f} mm): {seg}")
    else:
        spline_array_back = []
        for (px, py) in back_pts:
            spline_array_back.extend([py * scale, px * scale, 0.0])
        pt_data = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, spline_array_back)
        seg = swModel.SketchManager.CreateSpline2(pt_data, True)
        print(f"  Back surface spline ({num_points} pts): {seg}")

    # --- Create construction centerline (revolve axis along sketch Y = optical axis) ---
    seg = swModel.SketchManager.CreateCenterLine(
        0.0, -x_extent * scale, 0.0,
        0.0, x_extent * scale, 0.0
    )
    print(f"  Centerline created: {seg} (extent={x_extent:.4f} mm)")

    # Disable direct database mode
    swModel.SketchManager.AddToDB = False

    # --- Exit sketch ---
    swModel.SketchManager.InsertSketch(True)
    print("  Sketch closed.")

    # --- Perform revolve ---
    swModel.ClearSelection2(True)
    boolstatus = swModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 0, pythoncom.Nothing, 0)
    if not boolstatus:
        feat = swModel.FeatureByPositionReverse(0)
        if feat is not None:
            feat.Select2(False, 0)
            boolstatus = True
    print(f"  Sketch selected for revolve: {boolstatus}")

    revolve_angle = 2.0 * math.pi
    swFeat = swModel.FeatureManager.FeatureRevolve2(
        True, True, False, False, False, False,
        0, 0, revolve_angle, 0,
        False, False, 0, 0,
        0, 0, 0,
        True, True, True
    )

    if swFeat is not None:
        print("  Revolve feature created successfully!")
    else:
        print("  WARNING: Revolve feature may have failed. Check SolidWorks for errors.")
        print("  You may need to manually select the profile and axis, then revolve.")

    # --- Rebuild and zoom to fit ---
    try:
        swModel.ForceRebuild3(True)
    except Exception:
        pass
    swModel.ViewZoomtofit2()
    print("\n  Done! Lens solid model created in SolidWorks.")
    print(f"  Lens diameter: {2 * semi_diam:.4f} mm")
    print(f"  Center thickness: {thickness:.4f} mm")


# =============================================================================
# CLI
# =============================================================================
def parse_args():
    parser = argparse.ArgumentParser(
        description="Convert a Zemax OpticStudio lens to a SolidWorks revolved solid.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  %(prog)s MyLens.zmx --front 2 --back 3
  %(prog)s MyLens.zmx --front 7 --back 8 --points 200
  %(prog)s MyLens.zmx --front 2 --back 3 --template "C:\\Templates\\Part(mm).prtdot"

Configuration file:
  Place a 'zemax_to_solidworks.json' next to this script or in your home
  directory to set persistent defaults (template path, point count, etc.).
""",
    )
    parser.add_argument(
        "lens_file",
        help="Path to the Zemax .zmx lens file",
    )
    parser.add_argument(
        "--front", type=int, default=None,
        help="Front lens surface number in the LDE (default: from config or 2)",
    )
    parser.add_argument(
        "--back", type=int, default=None,
        help="Back lens surface number in the LDE (default: from config or 3)",
    )
    parser.add_argument(
        "--points", type=int, default=None,
        help="Number of sample points per surface (default: from config or 150)",
    )
    parser.add_argument(
        "--template", type=str, default=None,
        help="Path to SolidWorks .prtdot part template (default: auto-detect)",
    )
    return parser.parse_args()


# =============================================================================
# Main
# =============================================================================
def main():
    args = parse_args()
    config = load_config()

    # Merge CLI args over config defaults
    lens_file = os.path.abspath(args.lens_file)
    surf_front = args.front or config.get("default_front_surface", 2)
    surf_back = args.back or config.get("default_back_surface", 3)
    num_points = args.points or config.get("num_points", 150)
    if args.template:
        config["sw_template"] = args.template

    print(f"  Lens file: {lens_file}")
    print(f"  Surfaces:  front={surf_front}, back={surf_back}")
    print(f"  Points:    {num_points}")
    print()

    # Phase 1: Extract parameters from Zemax
    surfaces = extract_lens_params(lens_file, surf_front, surf_back)

    # Phase 2: Compute the half-profile
    profile_points, semi_diam, thickness, profile_info = compute_profile_points(
        surfaces, surf_front, surf_back, num_points
    )

    # Phase 3: Build the solid in SolidWorks
    build_solidworks_model(profile_points, semi_diam, thickness, num_points, config,
                           surfaces, surf_front, surf_back, profile_info)


if __name__ == '__main__':
    main()

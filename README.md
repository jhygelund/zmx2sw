# Zemax Lens → SolidWorks Solid Model

Automatically converts lens surfaces from a Zemax OpticStudio `.zmx` file into a revolved solid body in SolidWorks — bridging optical design and mechanical CAD without manual re-entry of surface geometry.

## Why This Exists

Optical designers define lens surfaces in Zemax (radii, conics, aspheric coefficients), but mechanical engineers need those same surfaces as solid models in SolidWorks for tolerancing, packaging, and assembly. Manually recreating aspheric profiles from coefficient tables is tedious and error-prone. This script automates the full pipeline in seconds.

## How It Works

The script operates in three phases:

### Phase 1: Extract Lens Parameters (ZOS-API)

Connects to OpticStudio in **headless mode** via the ZOS-API .NET interface (located automatically from the Windows registry). OpticStudio does not need to be running — the script launches a background instance, reads the lens data, and closes it automatically. Reads the specified front/back surface pair from the Lens Data Editor:

- Radius of curvature
- Conic constant
- Even asphere coefficients (Par1–Par8 → r², r⁴, ... r¹⁶)
- Clear semi-diameter
- Center thickness (axial distance between surfaces)

### Phase 2: Compute Profile Geometry

Generates the closed half-profile of the lens cross-section:

- **Spherical surfaces** (conic = 0, no aspheric terms): stored as exact geometry for arc representation
- **Aspheric/conic surfaces**: sampled as dense point arrays using the standard even asphere sag equation:

$$z(r) = \frac{c \cdot r^2}{1 + \sqrt{1 - (1+k)c^2 r^2}} + \sum_{i} \alpha_i \cdot r^{2i}$$

Each surface is sampled only to its own clear semi-diameter. If one surface has a smaller aperture than the other, a radial flat (horizontal line) extends it to the lens OD before the vertical edge wall connects the two sides.

### Phase 3: Build SolidWorks Model (COM API)

Connects to the running SolidWorks instance via COM automation and:

1. Creates a new part document (template auto-detected or user-specified)
2. Opens a sketch on the Right (YZ) Plane
3. Draws the half-profile:
   - **Plano surfaces** → straight line
   - **Spherical surfaces** → exact circular arc (`CreateArc`)
   - **Aspheric/conic surfaces** → spline through sampled points (`CreateSpline2`)
   - Flat annulus lines and edge wall as needed
4. Adds a construction centerline along the optical axis
5. Revolves 360° to produce the solid body

## Requirements

| Dependency | Purpose |
|-----------|---------|
| Windows 10/11 | COM and .NET interop required |
| Python 3.8+ | Tested with 3.10 |
| [pythonnet](https://pypi.org/project/pythonnet/) (`clr`) | .NET bridge to ZOS-API |
| [pywin32](https://pypi.org/project/pywin32/) (`win32com`) | COM automation for SolidWorks |
| [NumPy](https://pypi.org/project/numpy/) | Sag computation |
| Zemax OpticStudio | Installed with valid ZOS-API license (does not need to be running — used in headless mode) |
| SolidWorks | Must be running before script execution |

### Install Python Dependencies

```
pip install pythonnet pywin32 numpy
```

## Usage

```powershell
py zemax_to_solidworks_public.py <lens_file.zmx> --front <N> --back <M>
```

### Arguments

| Argument | Description |
|----------|-------------|
| `lens_file` | Path to the `.zmx` file (required) |
| `--front N` | Front surface number in the LDE (default: 2) |
| `--back M` | Back surface number in the LDE (default: 3) |
| `--points P` | Sample points per surface for splines (default: 150) |
| `--template` | Path to a SolidWorks `.prtdot` part template |

### Examples

```powershell
# Simple doublet element (surfaces 2 and 3)
py zemax_to_solidworks_public.py "C:\Zemax\Designs\MyDoublet.zmx" --front 2 --back 3

# Third element in a mobile phone lens (surfaces 7 and 8)
py zemax_to_solidworks_public.py "C:\Zemax\Designs\MobileLens.zmx" --front 7 --back 8 --points 200

# Specify a corporate part template
py zemax_to_solidworks_public.py MyLens.zmx --front 2 --back 3 --template "D:\Templates\Part(mm).prtdot"
```

> **Note:** Quote paths containing spaces.

## Configuration File

Place a `zemax_to_solidworks.json` file next to the script or in your home directory (`%USERPROFILE%`) to set persistent defaults:

```json
{
  "sw_template": "C:\\Path\\To\\Your\\Part(mm).prtdot",
  "num_points": 150,
  "default_front_surface": 2,
  "default_back_surface": 3
}
```

CLI arguments always override config file values.

## Surface Type Handling

| Surface Type | Sketch Entity | Geometry |
|--------------|---------------|----------|
| Plano (R = ∞) | Line | Exact |
| Standard sphere (conic = 0, no aspheric terms) | Circular arc | Exact |
| Even Asphere / Conic | Spline through N points | Approximation (increase `--points` for tighter fit) |

## Limitations

- **Single element at a time** — run once per lens element (surface pair). For a multi-element system, run the script multiple times with different `--front`/`--back` values.
- **Even Asphere only** — does not handle odd aspheres, Zernike, Q-type, or freeform surface types. Standard and Even Asphere cover the vast majority of rotationally symmetric optics.
- **Rotationally symmetric** — the revolve assumes axial symmetry. Torics, cylinders, and off-axis surfaces are not supported.
- **No material/coating data** — only geometry is transferred. Material assignment must be done in SolidWorks.
- **SolidWorks must be running** — the script attaches to an existing session via `GetActiveObject`.

## Troubleshooting

| Issue | Fix |
|-------|-----|
| `python was not found` | Use `py` launcher instead of `python`, or fix App Execution Aliases in Windows Settings |
| `Unable to locate Zemax OpticStudio` | Ensure OpticStudio is installed and the `HKCU\Software\Zemax\ZemaxRoot` registry key exists |
| `License is not valid for ZOSAPI use` | Your OpticStudio license must include API access (Professional or Premium) |
| `Failed to create new part document` | Check that SolidWorks is running and not blocked by a dialog |
| `Could not find a SolidWorks part template` | Set `sw_template` in your config file or pass `--template` |
| Revolve fails | Usually a profile gap — check that front/back surface numbers are adjacent in the LDE |

## License

MIT

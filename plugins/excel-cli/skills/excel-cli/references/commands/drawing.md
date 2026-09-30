### drawing

Worksheet drawing objects and sparklines

**Actions:** `list-objects`, `get-object`, `add-image`, `add-shape`, `add-text-box`, `add-connector`, `add-form-control`, `update-object`, `delete-object`, `list-sparklines`, `get-sparkline`, `add-sparkline`, `update-sparkline`, `delete-sparkline`

| Parameter | Description |
|-----------|-------------|
| `--session` | Session ID from 'session open' command |
| `--sheet` | Worksheet containing the drawing objects or sparklines (required) |
| `--object-name` | Existing drawing object's name, as returned by list-objects (required for: get-object, update-object, delete-object) (valid for: get-object, update-object, delete-object) |
| `--image-path` | Full path to a readable local image file (required for: add-image) (valid for: add-image) |
| `--name` | Optional name for the new drawing object (valid for: add-image, add-shape, add-text-box, add-connector, add-form-control) |
| `--left` | Left position in points from the worksheet edge (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--top` | Top position in points from the worksheet edge (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--width` | Object width in points (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--height` | Object height in points (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--lock-aspect-ratio` | Keep the image's aspect ratio when resizing (valid for: add-image) |
| `--shape-type` | AutoShape type, such as Rectangle, Oval, or a supported arrow/flowchart shape (valid for: add-shape) |
| `--text` | Text displayed by the shape, text box, or Forms control (required for: add-text-box) (valid for: add-shape, add-text-box, add-form-control, update-object) |
| `--fill-color` | Fill color as #RRGGBB (valid for: add-shape, add-text-box, update-object) |
| `--line-color` | Outline, connector, or sparkline color as #RRGGBB (valid for: add-shape, add-text-box, add-connector, update-object, add-sparkline, update-sparkline) |
| `--line-weight` | Line thickness in points (valid for: add-shape, add-connector, update-object) |
| `--font-size` | Text size in points (valid for: add-text-box, update-object) |
| `--font-color` | Text color as #RRGGBB (valid for: add-text-box, update-object) |
| `--connector-type` | Connector geometry: Straight, Elbow, or Curved (valid for: add-connector) |
| `--begin-x` | Starting horizontal position in points (valid for: add-connector) |
| `--begin-y` | Starting vertical position in points (valid for: add-connector) |
| `--end-x` | Ending horizontal position in points (valid for: add-connector) |
| `--end-y` | Ending vertical position in points (valid for: add-connector) |
| `--control-type` | Worksheet Forms control type; ActiveX/OLE controls are excluded (valid for: add-form-control) |
| `--linked-cell` | Cell binding for CheckBox, DropDown, ListBox, OptionButton, ScrollBar, or Spinner (valid for: add-form-control, update-object) |
| `--input-range` | Cell range supplying items to a DropDown or ListBox (valid for: add-form-control, update-object) |
| `--new-name` | New name for the object; omit to keep its name (valid for: update-object) |
| `--rotation` | Rotation angle in degrees (valid for: update-object) |
| `--visible` | Show or hide the object; omit to leave unchanged (valid for: update-object) |
| `--locked` | Lock the object; effective when worksheet protection is enabled (valid for: update-object) |
| `--placement` | Cell anchoring: 1=move and size, 2=move only, 3=free floating (valid for: update-object) |
| `--alternative-text` | Accessible description of the object (valid for: update-object) |
| `--location-range` | Cell or range displaying the sparkline group (required for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) (valid for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) |
| `--source-range` | Cell range supplying the sparkline data (required for: add-sparkline) (valid for: add-sparkline, update-sparkline) |
| `--sparkline-type` | Sparkline type: Line, Column, or WinLoss (valid for: add-sparkline, update-sparkline) |
| `--show-markers` | Display markers on line sparklines (valid for: add-sparkline, update-sparkline) |
| `--output` | Write output to file instead of stdout. For image results, decodes and saves as binary file |

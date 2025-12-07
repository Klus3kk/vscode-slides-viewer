# Changelog

### v0.0.1 

#### Added

* Support for **PPTX** (text, images, shapes, backgrounds)
* Support for **PPT** with partial parsing (text blocks + basic background)
* Support for **ODP** with strong rendering (text, images, tables, backgrounds)
* Support for **KEY** with basic APXL parsing (text, images, simple shapes)
* Custom renderers written fully in JavaScript: no native dependencies
* Slide navigation toolbar (next / prev, zoom in/out, reset zoom)
* Automatic TIFF -> PNG conversion for preview
* Internal fallbacks for unsupported objects
* Webview-based viewer with scaling and keyboard navigation

#### Known limitations

* EMF/WMF content (charts, diagrams) not supported in PPT/PPTX
* Complex grouped shapes may flatten incorrectly
* KEY layout reconstruction incomplete
* Fonts substituted by browser

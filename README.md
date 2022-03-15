## VB6-SVG provides full SVG and SVGZ support for classic VB6 projects

It does this thanks to resvg, a comprehensive portable SVG library by Yevhenii Reizner:

https://github.com/RazrFalcon/resvg

resvg is available under an MPL-2 license.  Please see resvg-LICENSE.md for full details.

### System requirements

This project is 100% portable.  You simply need to ship `resvg.dll` alongside your application (perhaps in an app subfolder).

This project has been tested on Win 7, 10, and 11.  It theoretically supports Windows Vista but this has not been manually verified.

This project will not work on Windows XP.  This is due to limitations building resvg with XP support.

### How to use VB6-SVG

VB6-SVG has three mandatory components.  All three must be added to your VB6 project:

1. `resvg.dll`
2. `svgSupport.bas`
3. `svgImage.cls`

`svgSupport` handles resvg initialization, shutdown, and a bunch of associated resource management and GDI interop.

`svgImage` is a lightweight convenience class for managing individual SVG image instances.  Create instances of this class by calling `svgSupport.LoadSVG_FromFile()`.

Adding SVG support to your own VB6 projects is simple:

1. Ensure you have read and understood both `LICENSE.md` (for the VB6 code) and `resvg-LICENSE.md` (for resvg).
2. Ship `resvg.dll` with your app.  This is a traditional DLL, not an ActiveX one, so you do not need to register it.  Just make sure it is available in a predictable location.
3. Somewhere in your VB6 project initialization, add one line of code:

`svgSupport.StartSVGSupport "C:\[path-to-resvg-folder]\resvg.dll"`

That line of code will initialize resvg and prepare a bunch of SVG-related resources.

4. Create (unlimited) `svgImage` instances by calling:

`svgSupport.LoadSVG_FromFile()`

Each `svgImage` instance manages a single SVG image.  `svgImage` stores a parsed SVG "tree", allowing you to render the SVG over-and-over at whatever position(s), size(s), and opacities you desire.  You can query individual instances for their default width/height, or draw them at whatever width/height you want using the `DrawSVGtoDC()` function.  As you'd expect for vector images, resizing and painting is always non-destructive.

5. `svgImage` instances manage their own resources.  You do not need to manage them manually, with one exception (see (6), below).

6. Before your program exits, ensure all `svgImage` instances have gone out of scope (or been manually freed), then add one line of code to your program shutdown process:

`svgSupport.StopSVGSupport`

This will free all shared GDI and SVG resources used by the project, then manually unload resvg itself.  

As you can imagine, if you are using module- or global- (*ugh*) `svgImage` class instances, they need to be freed **before** calling `StopSVGSupport`, because once resvg is released, SVG management is over.

7. (Optional) this project must do some manual image processing to allow painting SVGs to arbitrary Windows DCs.  Performance will be improve significantly if you compile your project to native code with the `Remove Array Bounds Checks` optimization enabled.  Please do this.

8. That's it!  If you encounter any bugs or unexpected behavior, please [file an issue at GitHub](https://github.com/tannerhelland/vb6-svg/issues).

### Licensing

The VB6 portion of this project is available under a Simplified BSD license.  Full details are provided in LICENSE.md.

resvg is available under an MPL-2 license.  Full details are provided in resvg-LICENSE.md.

Many thanks to [Yevhenii Reizner](https://github.com/RazrFalcon) for his work on resvg.
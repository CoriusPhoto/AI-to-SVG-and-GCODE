# AI-to-SVG-and-GCODE
This script was created to ease the process of creating packaging mock-up with laser engraver from a diecut created with Adobe Illustrator. It works paired
with the template Corius_SVG_GCODE_for_LASERGRBL.ait
The script addresses the issue of Illustrator natively exported SVG files not always compatible with LaserGRBL. The SVG generated with this script are only
converting path stroke (no text, no fill, no stroke effect like dash, no color accuracy) with the only purpose to import them into LaserGRBL (they may work with
other GRBL software though).
It also export directly into GCODE files supporting parameters from Illustrator file like laser power, feed speed and number of pass.
The scipt was developped in Javasrcript, with Adobe ExtendScript Toolkit CC.
Usage principle :
The script will use the paths from the 3 main layers “CUT”, “FOLD” and “TEST” to generate SVG files as well as GCODE files (.nc file extension) which can both
be loaded into LaserGRBL software. The main GCODE parameters can be changed directly from the Illustrator template on the “GCODE_PARAMS” layer.

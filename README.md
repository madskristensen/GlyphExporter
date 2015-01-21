## Glyph Exporter for Visual Studio

[![Build status](https://ci.appveyor.com/api/projects/status/3uw49mn555ac8kq6?svg=true)](https://ci.appveyor.com/project/madskristensen/glyphexporter)

Download from the 
[Visual Studio Gallery](https://visualstudiogallery.msdn.microsoft.com/b6cb739c-359d-4cce-a06d-9c75fe55dc51)
or get the
[nightly build](https://ci.appveyor.com/project/madskristensen/glyphexporter/build/artifacts)

This will export the various image icons that Visual Studio uses. 

The icons comes from 2 sources:

1. IGlyphService
2. IVsImageService

Any image icons present in those services will be exported to disk.

By going to **Tools -> Export Glyphs** you will be prompted to select 
a folder on disk where the images should be exported into. 

The images can be seen here http://glyphlist.azurewebsites.net/
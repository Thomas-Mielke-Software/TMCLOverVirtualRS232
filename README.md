# TMCLOverVirtualRS232 #

C# class library for Trinamic's TMCL serial interface for controlling TMCM stepper modules via a serial-to-rs485 converter.

### How to use ###

Currently binds to .NET Framework 4.6.1 but should easily be adaptable to work with .NET 6 or as .NET Standard library.

TMCLOverVirtualRS232Test is a VB console sample project, showing how to use the library.

### Credits ###

* Wolfgang Kurz, on whose work-in-progress code this library is based on 
* Alan Pich & Christian Weickhmann, who wrote a [Python library](https://github.com/NativeDesign/python-tmcl) for the same purpose and from which I got some inspiration to finish TMCLOverVirtualRS232
* MyCount GmbH for letting me open-source this

### License ###

All code released under LGPL 3.0. This basically means, you may release any software together with this library, but if you change the code of this library, you are obliged to release it under the same or a compatible license (see [here](LICENSE.md) for details). Pull requests welcome!

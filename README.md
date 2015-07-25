# GCAPlugins
A project with plugins for GCA5 - exporting, printing, etc.

Currently, the only plugin exports the character to a text format suitable for pasting into a MediaWiki site. It has some configuration options. Uses no special extensions or plugins on Mediawiki.

DLL and configuration XML goes into your Documents\GURPS Character Assistant 5\plugins\Somesubfolder\ folder. Somesubfolder is arbitrary, up to you.

Project build action currently includes copying the DLL, but not the plugin, to plugins\test automatically on build. Make sure your GCA5 isn't running at the time or the copy will fail and you'll get a build error.

This project is licensed under the MIT license, as follows:

The MIT License(MIT)
Copyright(c) 2015 Emily Smirle
Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in 
the Software without restriction, including without limitation the rights to 
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies 
of the Software, and to permit persons to whom the Software is furnished to do 
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.

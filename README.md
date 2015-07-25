# GCAPlugins
A project with plugins for GCA5 - exporting, printing, etc.

Currently, the only plugin exports the character to a text format suitable for pasting into a MediaWiki site. It has some configuration options. Uses no special extensions or plugins on Mediawiki.

DLL and configuration XML goes into your Documents\GURPS Character Assistant 5\plugins\Somesubfolder\ folder. Somesubfolder is arbitrary, up to you.

Project build action currently includes copying the DLL, but not the plugin, to plugins\test automatically on build. Make sure your GCA5 isn't running at the time or the copy will fail and you'll get a build error.

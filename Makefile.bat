@echo off
REM  This file is part of EAS-4-TbSync.
REM 
REM This Source Code Form is subject to the terms of the Mozilla Public
REM License, v. 2.0. If a copy of the MPL was not distributed with this
REM file, You can obtain one at http://mozilla.org/MPL/2.0/.

del EAS-4-TbSync.xpi
"C:\Program Files\7-Zip\7zG.exe" a -tzip EAS-4-TbSync.xpi content locale skin chrome.manifest install.rdf LICENSE README.md bootstrap.js



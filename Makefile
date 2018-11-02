# This file is part of EAS-4-TbSync.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

.PHONY: build

build:
	# Sync all files except git technical files
	# and screenshots
	zip ../EAS-4-TbSync.xpi -FS -R '*' \
		-x 'screenshots' '.git/*' '.gitignore'


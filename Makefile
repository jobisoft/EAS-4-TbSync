# This file is part of EAS-4-TbSync.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

PACKAGE_NAME=EAS-4-TbSync

ARCHIVE_NAME=$(PACKAGE_NAME).xpi

PACKAGE_FILES= \
	content \
	_locales \
	manifest.json \
	CONTRIBUTORS.md \
	LICENSE README.md \
	background.js

all: clean $(PACKAGE_FILES)
	zip -r $(ARCHIVE_NAME) $(PACKAGE_FILES)

clean:
	rm -f $(ARCHIVE_NAME)

.PHONY: clean

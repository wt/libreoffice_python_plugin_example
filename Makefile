FILES = \
	META-INF/manifest.xml \
	Addons.xcu \
	Addons.py

example-wt.oxt: $(FILES)
	zip -r $@ $^ pythonpath -x '*.pyc'

.PHONY: clean
clean:
	rm -f example-wt.oxt

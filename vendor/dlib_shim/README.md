# dlib shim

This local package exists so `pip install -r requirements.txt` can satisfy
`face-recognition`'s `dlib>=19.7` dependency on Python 3.12 without trying to
compile `dlib` from source on Windows.

It installs the prebuilt `dlib-bin==20.0.1` wheel and exposes no Python
modules of its own.

# MathML.lisp

This is David Drysdale's mathml package for Maxima,
based on mactex.lisp.

# License: GPL

# Note to Developers

Here are some notes to assist developers who wish to make changes.

## Testing the package

To test the basic functioning of the package, run:

```shell
make check
```

To validate the generated XHTML files, use the W3C validator:

https://validator.w3.org/check

## Testsuite

The testsuite is located in `example.mac`.

The file `mathmltest.mac` contains a special printer for input and
output lines, using `mathml`. It also defines a customized version of
`batch`, called `mathml_batch`, that ensures that the Maxima file is
`batch`-ed and the output that is printed creates a valid XHTML 1.1
plus MathML 2.0 plus SVG 1.1 document.

The file `mathmltest.sh` is a template for the shell script that `make
check` executes for each lisp that is configured.

The shell script `mathmltest` provides a simple script to run a single
test with the default lisp and an arbitrary number of test files:

```shell
./mathmltest <file1>.mac [... <fileN>.mac]
```

At the time of writing (2026-03-16), both tests produce a document
that validates on the W3C validator.

# Documentation

User documentation is in mathml.texi.

(in-package :maxima)

(unless (member :mk-defsystem *features*) ($load "lisp-utils/defsystem.lisp"))

(load (combine-path *maxima-sharedir* "affine" "affine.system"))

(mk:compile-system "affine" :load-source-if-no-binary t)

;;; affine.lisp ends here

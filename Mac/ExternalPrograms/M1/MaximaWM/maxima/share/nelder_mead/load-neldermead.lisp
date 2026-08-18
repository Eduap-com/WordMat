(in-package #:maxima)

(unless (member :mk-defsystem *features*) ($load "lisp-utils/defsystem.lisp"))

(load (merge-pathnames (make-pathname :name "nelder_mead" :type "system")
		       *load-pathname*))

(mk:oos "nelder_mead" :compile)

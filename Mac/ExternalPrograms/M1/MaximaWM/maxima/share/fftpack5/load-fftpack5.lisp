(in-package #:maxima)

(unless (member :mk-defsystem *features*) ($load "lisp-utils/defsystem.lisp"))

(load (merge-pathnames (make-pathname :name "fftpack5" :type "system")
		       (maxima-load-pathname-directory)))

(mk:oos "fftpack5-interface" :compile)

(in-package #-gcl #:maxima #+gcl "MAXIMA")

($put '$draw 2 '$version)

#+ecl ($load "lisp-utils/defsystem.lisp")

(load (merge-pathnames (make-pathname :name "draw" :type "system") (maxima-load-pathname-directory)))

(mk:oos "draw" :compile)

(in-package #-gcl #:maxima #+GCL "MAXIMA")

#+ecl
($load "lisp-utils/defsystem.lisp")

(load (merge-pathnames (make-pathname :name "cobyla" :type "system") (maxima-load-pathname-directory)))

(mk:oos "cobyla-interface" :compile)

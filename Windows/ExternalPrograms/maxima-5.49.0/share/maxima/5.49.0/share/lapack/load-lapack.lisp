(in-package #:maxima)

#+nil
(progn
  (format t "path = ~A~%" (combine-path *maxima-sharedir* "lapack"))
  (format t "(maxima-load-pathname-directory) = ~A~%" (maxima-load-pathname-directory))
  (format t "sys = ~A~%" (merge-pathnames (make-pathname :name "lapack" :type "system") (maxima-load-pathname-directory))))

#+ecl ($load "lisp-utils/defsystem.lisp")

(load (merge-pathnames (make-pathname :name "lapack" :type "system") (maxima-load-pathname-directory)))

;; Maxima errored out when any lapack function was used which
;; most certainly was an ECL bug: Seems like the definition of the
;; MAXIMA package shadows the array symbol from the COMMON-LISP package.
;; Bugfix by Marius Gerbershagen:
#+ecl (in-package #:common-lisp)

#-abcl (mk:oos "lapack-interface" :compile)

#+abcl (require "asdf")
#+abcl (push (maxima-load-pathname-directory) asdf:*central-registry*)
#+abcl (asdf:operate 'asdf:load-source-op "lapack")

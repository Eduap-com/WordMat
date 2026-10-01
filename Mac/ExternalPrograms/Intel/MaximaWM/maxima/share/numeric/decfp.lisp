;; -*- Lisp -*-

(in-package :maxima)

(unless (member :mk-defsystem *features*) ($load "lisp-utils/defsystem.lisp"))

(mk:defsystem decimal-fp
  :source-pathname (maxima::maxima-load-pathname-directory)
  :binary-pathname (maxima::maxima-objdir "share" "numeric")
  :source-extension "lisp"
  :components
  ((:file "decfp-core")))

(#+cmu
 ext:without-package-locks
 #-cmu
 progn
 (mk:oos "decimal-fp" :compile))



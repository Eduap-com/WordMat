; -*- Lisp -*-
(in-package :maxima)

(defparameter *autoconf-prefix* "C:/maxima-5.47.0")
(defparameter *autoconf-exec_prefix* "C:/maxima-5.47.0")
(defparameter *autoconf-package* "maxima")
(defparameter *autoconf-version* "5.47.0")
(defparameter *autoconf-libdir* "C:/maxima-5.47.0/lib")
(defparameter *autoconf-libexecdir* "C:/maxima-5.47.0/libexec")
(defparameter *autoconf-datadir* "C:/maxima-5.47.0/share")
(defparameter *autoconf-infodir* "C:/maxima-5.47.0/share/info")
(defparameter *autoconf-host* "x86_64-w64-mingw32")
;; This variable is kept for backwards compatibility reasons:
;; We seem to be in the fortunate position that we sometimes need to check for windows.
;; But at least until dec 2015 we didn't need to check for a specific windows flavour.
(defparameter *autoconf-win32* "true")
(defparameter *autoconf-windows* "true")
(defparameter *autoconf-ld-flags* "")

;; This will be T if this was a lisp-only build
(defparameter *autoconf-lisp-only-build* (eq t 'nil))
 
(defparameter *maxima-source-root* "/tmp/maxima-5.47.0/crosscompile-windows/build/maxima-prefix/src/maxima")
(defparameter *maxima-default-layout-autotools* "true")

; -*- Lisp -*-
(in-package :maxima)

(defparameter *autoconf-variables-set* "@variables_set@")
(defparameter *autoconf-prefix* "c:/maxima-sbcl")
(defparameter *autoconf-exec_prefix* "c:/maxima-sbcl")
(defparameter *autoconf-package* "maxima")
(defparameter *autoconf-version* "5.38.1")
(defparameter *autoconf-libdir* "c:/maxima-sbcl/lib")
(defparameter *autoconf-libexecdir* "c:/maxima-sbcl/libexec")
(defparameter *autoconf-datadir* "c:/maxima-sbcl/share")
(defparameter *autoconf-infodir* "c:/maxima-sbcl/share/info")
(defparameter *autoconf-host* "i686-pc-mingw32")
;; This variable is kept for backwards compatibiliy reasons:
;; We seem to be in the fortunate position that we sometimes need to check for windows.
;; But at least until dec 2015 we didn't need to check for a specific windows flavour.
(defparameter *autoconf-win32* "true")
(defparameter *autoconf-windows* "true")
(defparameter *autoconf-ld-flags* "")
 
(defparameter *maxima-source-root* "/C/Users/ProfesorToshiba/Desktop/Andrej/maxima-5.38.1")
(defparameter *maxima-default-layout-autotools* "true")

; -*- Lisp -*-
(in-package :maxima)

(defparameter *autoconf-prefix* "/opt/local")
(defparameter *autoconf-exec_prefix* "/opt/local")
(defparameter *autoconf-package* "maxima")
(defparameter *autoconf-version* "5.45.1")
(defparameter *autoconf-libdir* "/opt/local/lib")
(defparameter *autoconf-libexecdir* "/opt/local/libexec")
(defparameter *autoconf-datadir* "/opt/local/share")
(defparameter *autoconf-infodir* "/opt/local/share/info")
(defparameter *autoconf-host* "arm-apple-darwin20.6.0")
;; This variable is kept for backwards compatibiliy reasons:
;; We seem to be in the fortunate position that we sometimes need to check for windows.
;; But at least until dec 2015 we didn't need to check for a specific windows flavour.
(defparameter *autoconf-win32* "false")
(defparameter *autoconf-windows* "false")
(defparameter *autoconf-ld-flags* "-L/opt/local/lib -Wl,-headerpad_max_install_names -Wl,-syslibroot,/Applications/Xcode.app/Contents/Developer/Platforms/MacOSX.platform/Developer/SDKs/MacOSX11.3.sdk -arch arm64")

;; This will be T if this was a lisp-only build
(defparameter *autoconf-lisp-only-build* (eq t 'nil))
 
(defparameter *maxima-source-root* "/opt/local/var/macports/build/_opt_local_var_macports_sources_rsync.macports.org_macports_release_tarballs_ports_math_maxima/maxima/work/maxima-5.45.1")
(defparameter *maxima-default-layout-autotools* "true")

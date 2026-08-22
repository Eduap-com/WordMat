;;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;

;;;; Setup the mapping from the Maxima 'flonum float type to a CL float type.
;;;;
;;;; Add :flonum-long to *features* if you want flonum to be a
;;;; long-float.  Or add :flonum-double-double if you want flonum to
;;;; be a double-double (currently only for CMUCL).  Otherwise, you
;;;; get double-float as the flonum type.
;;;;
(in-package "MAXIMA")

;;;; Default double-float flonum.
(eval-when (:compile-toplevel :load-toplevel :execute)
  (setq *read-default-float-format* 'double-float))

#-(or flonum-long flonum-double-double)
(progn
;; Tell Lisp the float type for a 'flonum.
#-(or clisp abcl)
(deftype flonum (&optional low high)
  (cond (high
	 `(double-float ,low ,high))
	(low
	 `(double-float ,low))
	(t
	 'double-float)))

;; Some versions of clisp and ABCL appear to be buggy: (coerce 1 'flonum)
;; signals an error.  So does (coerce 1 '(double-float 0d0)).  But
;; (coerce 1 'double-float) returns 1d0 as expected.  So for now, make
;; flonum be exactly the same as double-float, without bounds.
#+(or clisp abcl)
(deftype flonum (&optional low high)
  (declare (ignorable low high))
  'double-float)

(defconstant +most-positive-flonum+ most-positive-double-float)
(defconstant +most-negative-flonum+ most-negative-double-float)
(defconstant +least-positive-flonum+ least-positive-double-float)
(defconstant +least-negative-flonum+ least-negative-double-float)
(defconstant +flonum-epsilon+ double-float-epsilon)
(defconstant +least-positive-normalized-flonum+ least-positive-normalized-double-float)
(defconstant +least-negative-normalized-flonum+ least-negative-normalized-double-float)

(defconstant +flonum-exponent-marker+ #\D)
)

#+flonum-long
(progn
;;;; The Maxima 'flonum can be a CL 'long-float on the Scieneer CL or CLISP,
;;;; but should be the same as 'double-float on other CL implementations.

  (eval-when (:compile-toplevel :load-toplevel :execute)
    (setq *read-default-float-format* 'long-float))

;; Tell Lisp the float type for a 'flonum.
(deftype flonum (&optional low high)
  (cond (high
	 `(long-float ,low ,high))
	(low
	 `(long-float ,low))
	(t
	 'long-float)))

(defconstant +most-positive-flonum+ most-positive-long-float)
(defconstant +most-negative-flonum+ most-negative-long-float)
(defconstant +least-positive-flonum+ least-positive-long-float)
(defconstant +least-negative-flonum+ least-negative-long-float)
(defconstant +flonum-epsilon+ long-float-epsilon)
(defconstant +least-positive-normalized-flonum+ least-positive-normalized-long-float)
(defconstant +least-negative-normalized-flonum+ least-negative-normalized-long-float)

(defconstant +flonum-exponent-marker+ #\L)

)

#+flonum-double-double
(progn

;;;; The Maxima 'flonum can be a 'kernel:double-double-float on the CMU CL.

  (eval-when (:compile-toplevel :load-toplevel :execute)
    (setq *read-default-float-format* 'kernel:double-double-float))

;; Tell Lisp the float type for a 'flonum.
(deftype flonum (&optional low high)
  (cond (high
	 `(kernel:double-double-float ,low ,high))
	(low
	 `(kernel:double-double-float ,low))
	(t
	 'kernel:double-double-float)))

;; While double-double can represent number as up to
;; most-positive-double-float, it can't really do operations on them
;; due to the way multiplication and division are implemented.  (I
;; don't think there's any workaround for that.)
;;
;; So, the largest number that can be used is the float just less than
;; 2^1024/(1+2^27).  This is the number given here.
(defconstant most-positive-double-double-hi
  (scale-float (cl:float (1- 9007199187632128) 1d0) 944))

(defconstant +most-positive-flonum+ (cl:float most-positive-double-double-hi 1w0))
(defconstant +most-negative-flonum+ (cl:float (- most-positive-double-double-hi 1w0)))
(defconstant +least-positive-flonum+ (cl:float least-positive-double-float 1w0))
(defconstant +least-negative-flonum+ (cl:float least-negative-double-float 1w0))
;; This is an approximation to a double-double epsilon.  Due to the
;; way double-doubles are represented, epsilon is actually zero
;; because 1+x = 1 only when x is zero.  But double-doubles only have
;; 106 bits of precision, so we use that as epsilon.
(defconstant +flonum-epsilon+ (scale-float 1w0 -106))
(defconstant +least-positive-normalized-flonum+ (cl:float least-positive-normalized-double-float 1w0))
(defconstant +least-negative-normalized-flonum+ (cl:float least-negative-normalized-double-float 1w0))

(defconstant +flonum-exponent-marker+ #\W)

)


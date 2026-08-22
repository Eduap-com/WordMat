;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;
;;; This file is part of the Maxima computer algebra system
;;; (https://sourceforge.net/projects/maxima/)
;;;
;;; Maxima is copyrighted by its authors and licensed under the GNU
;;; General Public License.  See COPYING and AUTHORS for details.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

;;; Utilities for determining if numerical evaluation should be done.

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; The following functions test if numerical evaluation has to be done.
;;; The functions should help to test for numerical evaluation more consistent
;;; and without complicated conditional tests including more than one or two
;;; arguments.
;;;
;;; The functions take a list of arguments. All arguments have to be a CL or
;;; Maxima number. If all arguments are numbers we have two cases:
;;; 1. $numer is T we return T. The function has to be evaluated numerically.
;;; 2. One of the args is a float or a bigfloat. Evaluate numerically.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun float-or-rational-p (x)
  (or (floatp x) ($ratnump x)))

(defun bigfloat-or-number-p (x)
  (or ($bfloatp x) (numberp x) ($ratnump x)))

;; Is argument u a complex number with real and imagpart satisfying predicate ntypep?
(defun complex-number-p (u &optional (ntypep 'numberp))
  (let ((R 0) (I 0))
    (labels ((a1 (x) (cadr x))
             (a2 (x) (caddr x))
             (a3+ (x) (cdddr x))
             (N (x) (funcall ntypep x)) ; N
             (i (x) (and (eq x '$%i) (N 1))) ; %i
             (N+i (x) (and (null (a3+ x)) ; mplus test is precondition
                           (N (setq R (a1 x)))
                           (or (and (i (a2 x)) (setq I 1) t)
                               (and (mtimesp (a2 x)) (N*i (a2 x))))))
             (N*i (x) (and (null (a3+ x))               ; mtimes test is precondition
                           (N (setq I (a1 x)))
                           (eq (a2 x) '$%i))))
      (declare (inline a1 a2 a3+ N i N+i N*i))
      (cond ((N u) (values t u 0)) ;2.3
            ((atom u) (if (i u) (values t 0 1))) ;%i
            ((mplusp u) (if (N+i u) (values t R I))) ;N+%i, N+N*%i
            ((mtimesp u) (if (N*i u) (values t R I))) ;N*%i
            (t nil)))))

;;; Check for an integer or a float or bigfloat representation. When we
;;; have a float or bigfloat representation return the integer value.

(defun integer-representation-p (x)
  (let ((val nil))
    (cond ((integerp x) x)
          ((and (floatp x) (= 0 (nth-value 1 (truncate x))))
           (nth-value 0 (truncate x)))
          ((and ($bfloatp x) 
                (eq ($sign (sub (setq val ($truncate x)) x)) '$zero))
           val)
          (t nil))))

(defun complexify (x)
  ;; Convert a Lisp number to a maxima number
  (cond ((realp x) x)
	((complexp x) (add (realpart x) (mul '$%i (imagpart x))))
	(t (merror (intl:gettext "COMPLEXIFY: argument must be a Lisp real or complex number.~%COMPLEXIFY: found: ~:M") x))))
   
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Helper functions for Bigfloat numerical evaluation.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defun cmul (x y) ($rectform (mul x y)))

(defun cdiv (x y) ($rectform (div x y)))

(defun cpower (x y) ($rectform (power x y)))

;;; Test for numerically evaluation in float precision

(defun float-numerical-eval-p (&rest args)
  (let ((flag nil))
    (dolist (ll args)
      (when (not (float-or-rational-p ll)) 
        (return-from float-numerical-eval-p nil))
      (when (floatp ll) (setq flag t)))
    (if (or $numer flag) t nil)))

;;; Test for numerically evaluation in complex float precision

(defun complex-float-numerical-eval-p (&rest args)
  "Determine if ARGS consists of numerical values by determining if
  the real and imaginary parts of each arg are nuemrical (but not
  bigfloats).  A non-NIL result is returned if at least one of args is
  a floating-point value or if numer is true.  If the result is
  non-NIL, it is a list of the arguments reduced via COMPLEX-NUMBER-P"
  (let (flag values)
    (dolist (ll args)
      (multiple-value-bind (bool rll ill)
          (complex-number-p ll 'float-or-rational-p)
        (unless bool
          (return-from complex-float-numerical-eval-p nil))
        ;; Always save the result from complex-number-p.  But for backward
        ;; compatibility, only set the flag if any item is a float.
        (push (add rll (mul ill '$%i)) values)
        (setf flag (or flag (or (floatp rll) (floatp ill))))))
    (when (or $numer flag)
      ;; Return the values in the same order as the args!
      (nreverse values))))

;;; Test for numerically evaluation in bigfloat precision

(defun bigfloat-numerical-eval-p (&rest args)
  (let ((flag nil))
    (dolist (ll args)
      (when (not (bigfloat-or-number-p ll)) 
        (return-from bigfloat-numerical-eval-p nil))
      (when ($bfloatp ll)
	(setq flag t)))
    (if (or $numer flag) t nil)))

;;; Test for numerically evaluation in complex bigfloat precision

(defun complex-bigfloat-numerical-eval-p (&rest args)
  "Determine if ARGS consists of numerical values by determining if
  the real and imaginary parts of each arg are nuemrical (including
  bigfloats). A non-NIL result is returned if at least one of args is
  a floating-point value or if numer is true. If the result is
  non-NIL, it is a list of the arguments reduced via COMPLEX-NUMBER-P."

  (let (flag values)
    (dolist (ll args)
      (multiple-value-bind (bool rll ill)
          (complex-number-p ll 'bigfloat-or-number-p)
        (unless bool
          (return-from complex-bigfloat-numerical-eval-p nil))
	;; Always save the result from complex-number-p.  But for backward
	;; compatibility, only set the flag if any item is a bfloat.
	(push (add rll (mul ill '$%i)) values)
	(when (or ($bfloatp rll) ($bfloatp ill))
          (setf flag t))))
    (when (or $numer flag)
      ;; Return the values in the same order as the args!
      (nreverse values))))

;;; Test for numerical evaluation in any precision, real or complex.
(defun numerical-eval-p (&rest args)
  (or (apply 'float-numerical-eval-p args)
      (apply 'complex-float-numerical-eval-p args)
      (apply 'bigfloat-numerical-eval-p args)
      (apply 'complex-bigfloat-numerical-eval-p args)))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;; Utilitye function to signal an error denoting a domain error for
;; the special functions.  ARGS should be appropriate for MERROR.
(defun simp-domain-error (&rest args)
  (if errorsw
      (throw 'errorsw t)
      (apply #'merror args)))



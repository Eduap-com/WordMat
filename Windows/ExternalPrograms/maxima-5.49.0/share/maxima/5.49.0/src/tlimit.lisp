;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     The data in this file contains enhancements.                   ;;;;;
;;;                                                                    ;;;;;
;;;  Copyright (c) 1984,1987 by William Schelter,University of Texas   ;;;;;
;;;     All rights reserved                                            ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     (c) Copyright 1980 Massachusetts Institute of Technology         ;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

(declare-top (special taylored *getsignl-asksign-ok*))

(macsyma-module tlimit)

(load-macsyma-macros rzmac)

;; For limits toward `inf`, assume that the limit variable exceeds `*large-positive-number*`
(defmvar *large-positive-number* 4398046511104) ; 2^42 for no particular reason
;; TOP LEVEL FUNCTION(S): $TLIMIT $TLDEFINT

(defmfun $tlimit (&rest args)
  (let ((limit-using-taylor t))
    (declare (special limit-using-taylor))
    (apply #'$limit args)))

(defmfun $tldefint (exp var ll ul)
  (let ((limit-using-taylor t))
    (declare (special limit-using-taylor))
    ($ldefint exp var ll ul)))

;; Taylor cannot handle conjugate, ceiling, floor, unit_step, or signum 
;; expressions, so let's tell tlimit to *not* try. We also disallow 
;; expressions containing $ind.

;; We have subst([h=0,x=0], taylor(asin(x+h)-asin(x),h,0,1)) = %pi (see
;; bug #4416 limit of Newton quotient involving asin). This bug causes
;; trouble for tlimit((asin(x+h) - asin(x))/h,h,0). Until such bugs are
;; sorted, we will disallow tlimit from attempting limits involving 
;; acos and asin.
(defun tlimp (e x)	
  (or (and ($mapatom e) (not (eq e '$ind)) (not (eq e '$und)))
	  (and (consp e) 
	       (consp (car e)) 
		   (or (not (among x e)) 
		     (not (member (caar e) '($conjugate $ceiling $floor $unit_step %signum %acos %asin) :test #'eq)))
	       (or 
		      (known-ps (caar e)) 
			  (and (eq (caar e) 'mqapply) (known-ps (subfunname e)))
	          (member (caar e) (list 'mplus 'mtimes 'mexpt '%log))
			  (get (caar e) 'grad)
			  ($freeof x e))
		    (every #'(lambda (q) (tlimp q x)) (cdr e)))))

(defun logarc-atan2 (e)
"Use `logarc` to transform all occurrences of `%atan2` subexpression in `e`, but do not 
 transform other log-like functions."
  (cond (($mapatom e) e)
        ((eq (caar e) '%atan2)
    	    (logarc '%atan2 (list (logarc-atan2 (second e)) (logarc-atan2 (third e)))))
        (t (recur-apply #'logarc-atan2 e))))

(defun atan2-to-atan (e)
 "In the expression `e`, replace all subexpressions of the form atan2(y,x), where y is not explicitly equal to 
  zero, by 2*atan((sqrt(x^2+y^2) - x)/y).  The input `e` should be simplified--that way, the general simplifier 
  handles the error case of atan2(0,0) and many other cases as well. "
  (cond (($mapatom e) e)
        ((and (consp e) (eq '%atan2 (caar e))) 
          (let ((y (second e)) (x (third e)))
            (if (zerop1 y)
                e
                (mul 2 (ftake '%atan (div (sub (ftake 'mexpt (add (mul x x) (mul y y)) (div 1 2)) x) y))))))
        (($subvarp (mop e)) ;subscripted function
             (subfunmake
              (subfunname e)
              (subfunsubs e) ;don't map fun onto the operator subscripts
              (mapcar #'atan2-to-atan (subfunargs e)))) ; map onto the arguments
        (t (fapply (caar e) (mapcar #'atan2-to-atan (cdr e))))))

;; Dispatch Taylor, but recurse on the order until either the recursion
;; depth reaches 15 or the Taylor polynomial is nonzero. If Taylor 
;; fails to find a nonzero Taylor polynomial or the recursion depth 
;; exceeds the limit, return NIL.

;; This recursion on the order attempts to handle limits such as 
;; tlimit(2^n / n^5, n, inf) correctly.

;; We set up a reasonable environment for calling Taylor. When $taylor_logexpand 
;; is false, Taylor sometimes generates expressions that vanish but do not readily
;; simplify to zero. For example:
;;
;;   block([taylor_logexpand: false], subst(h=0, taylor(atan(x+h) - atan(x), h, 0, 1)));
;;
;; For the limit code, such expressions can cause errors. To avoid this, we set 
;; taylor_logexpand to true. Previously, this code set the value of $taylor_simplifier
;; to  #'(lambda (q) (sratsimp (extra-simp q)))), but that doesn't seem to be needed,
;; so this version doesn't set a value for this option variable.

;; There is no compelling reason to default the Taylor order to 
;; lhospitallim, but this is documented in the user documentation.

;; Since `taylor` fails on some `atan2` expressions, we convert `atan2` expressions to
;; log form. An example where `taylor` fails is `taylor(atan2(exp(x)-cos(x), -x*sin(x)),x,0,4)`.
;; Of course, we should fix `taylor`, but until that happens, we'll use this workaround.
(defun tlimit-taylor (e x pt n &optional (d 0))
  "Compute the Taylor series expansion of `e` at `pt` with respect to `x`. 
   If the expansion vanishes and the recursion depth `d` is less than 16, 
   retry with increased order. If recursion depth exceeds 16 or the expansion 
   fails, return nil; otherwise, return the Taylor expansion."
	(let ((ee) 
	      (silent-taylor-flag t) 
	      ($taylordepth 8)
		  ($radexpand nil)
		  ($taylor_logexpand t)
		  ($logexpand nil))
    
    (cond
      ((eq pt '$infinity) nil)
      (t
       (setq e (atan2-to-atan e))
       (setq ee (catch 'taylor-catch ($totaldisrep ($taylor e x pt n))))
       (cond
         ((and ee (not (eql ee 0))) ee)
         ;; Retry if Taylor returns zero and depth is less than 16
         ((and ee (< d 16))
          (tlimit-taylor e x pt (* 2 (max 1 n)) (1+ d)))
         (t nil))))))

;; Previously, when the Taylor series failed, there was code to decide
;; whether to call limit1 or simplimit. The choice depended on the last
;; argument to taylim (previously named *i*) and the main operator of the 
;; expression. This updated code eliminates that logic and always dispatches
;; limit1 when Maxima is unable to find the Taylor polynomial. As a result, 
;; the last argument of taylim is now unused (orphaned).
(defun taylim (e var val flag)
  "Attempt to compute the limit of `e` as `var` approaches `val` using a Taylor expansion.
  When the Taylor expansion fails, fall back to using `limit1`. The `flag` argument is unused."
  (declare (ignore flag))
  (let ((et nil))
    (when (tlimp e var)
      (setq e (stirling0 e))
      (setq et (tlimit-taylor e var (ridofab val) $lhospitallim 0)))
    (if et
        (let ((taylored t)) ; the special variable `taylored` prevents infinite looping
          (or (limit-sum-of-powers et var val)
              (limit et var val 'think)))
        (limit1 e var val))))

(defun power-of-x-p (e x)
 "Return true if `e` is a monomial of the form `x^q`, where `q` is a rational number. Also returns true if `e` is `x`.
  We allow `q` to be one. The case of `q=0` shouldn't happen unless the expression `e` is not simplified. The second
  argument `x` should be a symbol, but that condition is not checked."
   (or 
      (eq e x)
      (and (mexptp e)
           (eq x (second e))
           ($ratnump (third e)))))

(defun sum-of-powers-p (e x)
"Return true iff `e` is a linear combination of monomials of the form `x^q`, where `q` is a rational number. The
exponent `q` can be zero, so a term that is free of `x` is acceptable. The second argument `x` should be a symbol, 
but that condition is not checked."
  (cond ((or ($mapatom e) (freeof x e)) t)
        ((power-of-x-p e x))
        ((mtimesp e) (every #'(lambda (q) (or (freeof x q) (power-of-x-p q x))) (cdr e)))
        ((mplusp e) (every #'(lambda (q) (sum-of-powers-p q x)) (cdr e)))
        (t nil)))

(defun limit-sum-of-powers (e x pt)
  "If `e` is a linear combination of terms like `x^q`, where `q` is an explicit rational number, 
   return limit(e, x, pt); otherwise, return nil. The limit point `pt` must be one of `minf`, 
   `zerob`, `zeroa`, `0`, or `inf`."
  (let ((ee) (ll nil) (pk nil) (ck nil) (sgn))
    (cond
      ((sum-of-powers-p e x)
       (cond
         ((or (eq pt '$zeroa) (eq pt '$zerob) (eql pt 0))
          ;; Try direct substitution
          (let* (($errormsg nil) (ans (errcatch (maxima-substitute 0 x e))))
            (cond
              ((null ans)
               ;; direct substitution failed, transform limit point to inf or minf and try again
               (let ((cntx ($supcontext)) (g (gensym)))
                 (unwind-protect
                   (progn
                      (putprop g t 'internal) ;ask no questions about the gensym variable
                      (assume (ftake 'mgreaterp g *large-positive-number*))
                      (limit-sum-of-powers 
                          (resimplify (maxima-substitute (div 1 g) x e)) g (if (eq pt '$zeroa) '$inf '$minf)))
                    ($killcontext cntx)))) 
              (t
               ;; direct substitution succeeded--when the limit is zero, use zero-fixup
               (setq ans (car ans))
               (if (eq t (meqp ans 0))
                   (zero-fixup e x pt)
                   ans)))))

         ((or (eq pt '$minf) (eq pt '$inf))
          ;; Set ee to a list of the additive terms of e
          (setq ee (if (mplusp e) (cdr e) (list e)))
          ;; Replace each term ck*x^pk of the the list `ee` by `ck . pk` and push them in the list `ll`
          (dolist (ek ee)
              (setq pk (sratsimp (mul x (div (sdiff ek x) ek))))
              (setq ck (sratsimp (div ek (ftake 'mexpt x pk))))
              (push (cons ck pk) ll))
          ;; Sort the terms of `ll` and determine the highest power
          (setq ll (sort ll #'(lambda (a b) (eq t (mgrp a b))) :key #'cdr))
          (setq pk (cdr (first ll))) ;pk is now the largest power
          ;; Process leading coefficients
          (setq ll (mapcar #'(lambda (q) (if (eq t (meqp pk (cdr q))) (car q) 0)) ll))
          (setq ck (fapply 'mplus ll)) ;ck is now the leading coefficient
          (cond
            ((eql pk 0) ck)  ;; Leading term is ck x^0, so limit is ck
            ((eq t (mgrp pk 0)) ;; Leading term is ck x^pk where pk > 0
             (setq sgn (let ((*getsignl-asksign-ok* t)) (maybe-asksign ck)))
             (cond
               ((eq sgn '$neg) (if (eq pt '$inf) '$minf '$inf))
               ((eq sgn '$pos) (if (eq pt '$inf) '$inf '$minf))
               ((eq sgn '$imaginary) '$infinity)
               ((eq sgn '$complex) '$infinity)
               (t nil)))
            ((eq t (mgrp 0 pk)) ;leading term is ck x^pk where pk < 0
              ;; the limit is a zero, so dispatch `zero-fixup`
              (zero-fixup (mul ck (power x pk)) x pt))
            (t nil))))) ;unexpected case
      ;; not a sum of powers, return nil
      (t nil))))

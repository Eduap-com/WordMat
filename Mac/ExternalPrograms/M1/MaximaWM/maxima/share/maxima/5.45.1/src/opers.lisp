;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     The data in this file contains enhancments.                    ;;;;;
;;;                                                                    ;;;;;
;;;  Copyright (c) 1984,1987 by William Schelter,University of Texas   ;;;;;
;;;     All rights reserved                                            ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     (c) Copyright 1980 Massachusetts Institute of Technology         ;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

(macsyma-module opers)

;; This file is the run-time half of the OPERS package, an interface to the
;; Macsyma general representation simplifier.  When new expressions are being
;; created, the functions in this file or the macros in MOPERS should be called
;; rather than the entrypoints in SIMP such as SIMPLIFYA or SIMPLUS.  Many of
;; the functions in this file will do a pre-simplification to prevent
;; unnecessary consing. [Of course, this is really the "wrong" thing, since
;; knowledge about 0 being the additive identity of the reals is now
;; kept in two different places.]

;; The basic functions in the virtual interface are ADD, SUB, MUL, DIV, POWER,
;; NCMUL, NCPOWER, NEG, INV.  Each of these functions assume that their
;; arguments are simplified.  Some functions will have a "*" adjoined to the
;; end of the name (as in ADD*).  These do not assume that their arguments are
;; simplified.  In addition, there are a few entrypoints such as ADDN, MULN
;; which take a list of terms as a first argument, and a simplification flag as
;; the second argument.  The above functions are the only entrypoints to this
;; package.

;; The functions ADD2, ADD2*, MUL2, MUL2*, and MUL3 are for use internal to
;; this package and should not be called externally.  Note that MOPERS is
;; needed to compile this file.

;; Addition primitives.

(defun add2 (x y)
  (cond ((numberp x)
	 (cond ((numberp y) (+ x y))
               ((=0 x) y)
	       (t (simplifya `((mplus) ,x ,y) t))))
        ((=0 y) x)
	(t (simplifya `((mplus) ,x ,y) t))))

(defun add2* (x y)
  (cond
    ((and (numberp x) (numberp y)) (+ x y))
    ((=0 x) (simplifya y nil))
    ((=0 y) (simplifya x nil))
    (t (simplifya `((mplus) ,x ,y) nil))))

;; The first two cases in this cond shouldn't be needed, but exist
;; for compatibility with the old OPERS package.  The old ADDLIS
;; deleted zeros ahead of time.  Is this worth it?

(defun addn (terms simp-flag)
  (cond ((null terms) 0)
	(t (simplifya `((mplus) . ,terms) simp-flag))))

(declare-top (special $negdistrib))

(defun neg (x)
  (cond ((numberp x) (- x))
	(t (let (($negdistrib t))
	     (simplifya `((mtimes) -1 ,x) t)))))

(defun sub (x y)
  (cond
    ((and (numberp x) (numberp y)) (- x y))
    ((=0 y) x)
    ((=0 x) (neg y))
    (t (add x (neg y)))))

(defun sub* (x y)
  (cond
    ((and (numberp x) (numberp y)) (- x y))
    ((=0 y) x)
    ((=0 x) (neg y))
    (t
     (add (simplifya x nil) (mul -1 (simplifya y nil))))))

;; Multiplication primitives -- is it worthwhile to handle the 3-arg
;; case specially?  Don't simplify x*0 --> 0 since x could be non-scalar.

(defun mul2 (x y)
  (cond
    ((and (numberp x) (numberp y)) (* x y))
    ((=1 x) y)
    ((=1 y) x)
    (t (simplifya `((mtimes) ,x ,y) t))))

(defun mul2* (x y)
  (cond
    ((and (numberp x) (numberp y)) (* x y))
    ((=1 x) (simplifya y nil))
    ((=1 y) (simplifya x nil))
    (t (simplifya `((mtimes) ,x ,y) nil))))

(defun mul3 (x y z)
  (cond ((=1 x) (mul2 y z))
	((=1 y) (mul2 x z))
	((=1 z) (mul2 x y))
	(t (simplifya `((mtimes) ,x ,y ,z) t))))

;; The first two cases in this cond shouldn't be needed, but exist
;; for compatibility with the old OPERS package.  The old MULSLIS
;; deleted ones ahead of time.  Is this worth it?

(defun muln (factors simp-flag)
  (cond ((null factors) 1)
	((atom factors) factors)
	(t (simplifya `((mtimes) . ,factors) simp-flag))))

(defun div (x y)
  (if (=1 x)
      (inv y)
      (cond
        ((and (floatp x) (floatp y))
         (/ x y))
        ((and ($bfloatp x) ($bfloatp y))
         ;; Call BIGFLOATP to ensure that arguments have same precision.
         ;; Otherwise FPQUOTIENT could return a spurious value.
         (bcons (fpquotient (cdr (bigfloatp x)) (cdr (bigfloatp y)))))
        (t
          (mul x (inv y))))))

(defun div* (x y)
  (if (=1 x)
      (inv* y)
      (cond
        ((and (floatp x) (floatp y))
         (/ x y))
        ((and ($bfloatp x) ($bfloatp y))
         ;; Call BIGFLOATP to ensure that arguments have same precision.
         ;; Otherwise FPQUOTIENT could return a spurious value.
         (bcons (fpquotient (cdr (bigfloatp x)) (cdr (bigfloatp y)))))
        (t
          (mul (simplifya x nil) (inv* y))))))

(defun ncmul2 (x y)
  (simplifya `((mnctimes) ,x ,y) t))

(defun ncmuln (factors flag)
  (simplifya `((mnctimes) . ,factors) flag))

;; Exponentiation

;; Don't use BASE as a parameter name since it is special in MacLisp.

(defun power (*base power)
  (cond ((=1 power) *base)
	(t (simplifya `((mexpt) ,*base ,power) t))))

(defun power* (*base power)
  (cond ((=1 power) (simplifya *base nil))
	(t (simplifya `((mexpt) ,*base ,power) nil))))

(defun ncpower (x y)
  (cond ((=0 y) 1)
	((=1 y) x)
	(t (simplifya `((mncexpt) ,x ,y) t))))

;; [Add something for constructing equations here at some point.]

;; (ROOT X N) takes the Nth root of X.
;; Warning! Simplifier may give a complex expression back, starting from a
;; positive (evidently) real expression, viz. sqrt[(sinh-sin) / (sin-sinh)] or
;; something.

(defun root (x n)
  (cond ((=0 x) 0)
	((=1 x) 1)
	(t (simplifya `((mexpt) ,x ((rat simp) 1 ,n)) t))))

;; (Porm flag expr) is +expr if flag is true, and -expr
;; otherwise.  Morp is the opposite.  Names stand for "plus or minus"
;; and vice versa.

(defun porm (s x) (if s x (neg x)))
(defun morp (s x) (if s (neg x) x))

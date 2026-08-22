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

(macsyma-module mmacro)

;; Exported functions are MDEFMACRO, $MACROEXPAND, $MACROEXPAND1, MMACRO-APPLY
;;                        MMACROEXPANDED, MMACROEXPAND and MMACROEXPAND1


;; $MACROS declared in jpg;mlisp >


(defmvar $macroexpansion ()
  "Governs the expansion of Maxima Macros.  The following settings are
available:  FALSE means to re-expand the macro every time it gets called.
EXPAND means to remember the expansion for each individual call do that it 
won't have to be re-expanded every time the form is evaluated.  The form will 
still grind and display as if the expansion had not taken place.  DISPLACE
means to completely replace the form with the expansion.  This is more space
efficient than EXPAND but grinds and displays the expansion instead of the
call."
  modified-commands '($macroexpand)
  :setting-list      (() $expand $displace))


;;; LOCAL MACRO ;;;

(defmacro copy1cons (name) `(cons (car ,name) (cdr ,name)))

;;; DEFINING A MACRO ;;;

(defmspec mdefmacro (form) (setq form (cdr form))
	  (cond ((or (null (cdr form)) (cdddr form))
		 (merror (intl:gettext "macro definition: must have exactly two arguments; found: ~M")
			 `((mdefmacro) ,@form))
		 )
		(t (mdefmacro1 (car form) (cadr form)))))


(defun mdefmacro1 (fun body)
  (let ((name) (args))
    (cond ((or (atom fun)
	       (not (atom (caar fun)))                
	       (member 'array (cdar fun) :test #'eq)              
	       (mopp (setq name ($verbify (caar fun))))
	       (member name '($all $% $%% mqapply) :test #'eq))
	   (merror (intl:gettext "macro definition: illegal definition: ~M") ;ferret out all the
		   fun))		;  illegal forms
	  ((not (eq name (caar fun)))	;efficiency hack I guess
	   (rplaca (car fun) name)))	;  done in jpg;mlisp
    (setq args (cdr fun))		;  (in MDEFINE).
    (let ((dup (find-duplicate args :test #'eq :key #'mparam)))
      (when dup
        (merror (intl:gettext "macro definition: ~M occurs more than once in the parameter list") (mparam dup))))
    (mredef-check name)
    (do ((a args (cdr a)) (mlexprp))
	((null a)
	 (remove1 (ncons name) 'mexpr t $functions t) ;do all arg checking,
	 (cond (mlexprp (mputprop name t 'mlexprp)) ; then remove MEXPR defn
	       (t nil)))
      (cond ((mdefparam (car a)))
	    ((and (mdeflistp a)
		  (mdefparam (cadr (car a))))
	     (setq mlexprp t))
	    (t 
	     (merror (intl:gettext "macro definition: bad argument: ~M")
		     (car a)))))
    (remove-transl-fun-props name)
    (add2lnc `((,name) ,@args) $macros)
    (mputprop name (mdefine1 args body) 'mmacro)
     
    (cond ($translate (translate-and-eval-macsyma-expression
		       `((mdefmacro) ,fun ,body))))
    `((mdefmacro simp) ,fun ,body)))










;;; MACROEXPANDING FUNCTIONS ;;;


(defmspec $macroexpand (form) (setq form (cdr form))
	  (cond ((or (null form) (cdr form))
		 (merror (intl:gettext "macroexpand: must have exactly one argument; found: ~M")
			 `(($macroexpand) ,@form)))
		(t (mmacroexpand (car form)))))

(defmspec $macroexpand1 (form) (setq form (cdr form))
	  (cond ((or (null form) (cdr form))
		 (merror (intl:gettext "macroexpand1: must have exactly one argument; found: ~M")
			 `(($macroexpand1) ,@form)))
		(t (mmacroexpand1 (car form)))))


;;; SIMPLIFICATION ;;;

(defprop mdefmacro simpmdefmacro operators)

;; emulating simpmdef (for mdefine) in jm;simp
(defun simpmdefmacro (x ignored simp-flag)
  (declare (ignore ignored simp-flag))
  (cons '(mdefmacro simp) (cdr x)))


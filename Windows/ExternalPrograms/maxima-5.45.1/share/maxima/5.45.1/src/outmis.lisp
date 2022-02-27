;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     The data in this file contains enhancments.                    ;;;;;
;;;                                                                    ;;;;;
;;;  Copyright (c) 1984,1987 by William Schelter,University of Texas   ;;;;;
;;;     All rights reserved                                            ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

;;	** (c) Copyright 1982 Massachusetts Institute of Technology **

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;                                                                ;;;
;;;                Miscellaneous Out-of-core Files                 ;;;
;;;                                                                ;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(macsyma-module outmis)


(declare-top (special $exptisolate $labels $dispflag errorsw))

(defmvar $exptisolate nil)
(defmvar $isolate_wrt_times nil)

(defmfun $isolate (e *xvar)
  (iso1 e (getopr *xvar)))

(defun iso1 (e *xvar)
  (cond ((specrepp e) (iso1 (specdisrep e) *xvar))
	((and (free e 'mplus) (or (null $isolate_wrt_times) (free e 'mtimes))) e)
	((freeof *xvar e) (mgen2 e))
	((alike1 *xvar e) *xvar)
	((member (caar e) '(mplus mtimes) :test #'eq) (iso2 e *xvar))
	((eq (caar e) 'mexpt)
	 (cond ((null (atom (cadr e))) (list (car e) (iso1 (cadr e) *xvar) (caddr e)))
	       ((or (alike1 (cadr e) *xvar) (not $exptisolate)) e)
	       (t (let ((x ($rat (caddr e) *xvar)) (u 0) (h 0))
		    (setq u (ratdisrep ($ratnumer x)) x (ratdisrep ($ratdenom x)))
		    (if (not (equal x 1))
			(setq u ($multthru (list '(mexpt) x -1) u)))
		    (if (mplusp u)
			(setq u ($partition u *xvar) h (cadr u) u (caddr u)))
		    (setq u (power* (cadr e) (iso1 u *xvar)))
		    (cond ((not (equal h 0))
			   (mul2* (mgen2 (power* (cadr e) h)) u))
			  (t u))))))
	(t (cons (car e) (mapcar #'(lambda (e1) (iso1 e1 *xvar)) (cdr e))))))

(defun iso2 (e *xvar)
  (prog (hasit doesnt op)
     (setq op (ncons (caar e)))
     (do ((i (cdr e) (cdr i))) ((null i))
       (cond ((freeof *xvar (car i)) (setq doesnt (cons (car i) doesnt)))
	     (t (setq hasit (cons (iso1 (car i) *xvar) hasit)))))
     (cond ((null doesnt) (go ret))
	   ((and (null (cdr doesnt)) (atom (car doesnt))) (go ret))
	   ((prog2 (setq doesnt (simplify (cons op doesnt)))
		(and (free doesnt 'mplus)
		     (or (null $isolate_wrt_times)
			 (free doesnt 'mtimes)))))
	   (t (setq doesnt (mgen2 doesnt))))
     (setq doesnt (ncons doesnt))
     ret  (return (simplifya (cons op (nconc hasit doesnt)) nil))))

(defun mgen2 (h)
  (cond ((memsimilarl h (cdr $labels) (getlabcharn $linechar)))
	(t (setq h (displine h)) (and $dispflag (mterpri)) h)))

(defun memsimilarl (item list linechar)
  (cond ((null list) nil)
	((and (char= (getlabcharn (car list)) linechar)
	      (boundp (car list))
	      (memsimilar item (car list) (symbol-value (car list)))))
	(t (memsimilarl item (cdr list) linechar))))

(defun memsimilar (item1 item2 item2ev)
  (cond ((equal item2ev 0) nil)
	((alike1 item1 item2ev) item2)
	(t (let ((errorsw t) r)
	     (setq r (catch 'errorsw (div item2ev item1)))
	     (and (mnump r) (not (zerop1 r)) (div item2 r))))))

(defmfun $pickapart (x lev)
  (setq x (format1 x))
  (cond ((not (fixnump lev))
	 (merror (intl:gettext "pickapart: second argument must be an integer; found: ~M") lev))
	((or (atom x) (and (eq (caar x) 'mminus) (atom (cadr x)))) x)
	((= lev 0) (mgen2 x))
	((and (atom (cdr x)) (cdr x)) x)
	(t (cons (car x) (mapcar #'(lambda (y) ($pickapart y (1- lev))) (cdr x))))))

(defmfun $reveal (e lev)
  (setq e (format1 e))
  (if (and (fixnump lev) (plusp lev))
      (reveal e 1 lev)
      (merror (intl:gettext "reveal: second argument must be a positive integer; found: ~M") lev)))

(defun simple (x)
  (or (atom x) (member (caar x) '(rat bigfloat) :test #'eq)))

(defun reveal (e nn lev)
  (cond ((simple e) e)
	((= nn lev)
	 (cond ((eq (caar e) 'mplus) (cons '(|$Sum| simp) (ncons (length (cdr e)))))
	       ((eq (caar e) 'mtimes) (cons '(|$Product| simp) (ncons (length (cdr e)))))
	       ((eq (caar e) 'mexpt) '|$Expt|)
	       ((eq (caar e) 'mquotient) '|$Quotient|)
	       ((eq (caar e) 'mminus) '|$Negterm|)
	       ((eq (caar e) 'mlist)
	        (cons '(|$List| simp) (ncons (length (cdr e)))))
	       (t (getop (mop e)))))
	(t (let ((u (cond ((member 'simp (cdar e) :test #'eq) (car e))
			  (t (cons (caar e) (cons 'simp (cdar e))))))
		 (v (mapcar #'(lambda (x) (reveal (format1 x) (1+ nn) lev))
			    (margs e))))
	     (cond ((eq (caar e) 'mqapply) (cons u (cons (cadr e) v)))
		   ((eq (caar e) 'mplus) (cons u (nreverse v)))
		   (t (cons u v)))))))

(declare-top (special atvars munbound $props $gradefs $features opers
		      $contexts $activecontexts $aliases))

(defmspec $properties (x)
  (setq x (getopr (fexprcheck x)))
  (unless (or (symbolp x) (stringp x))
    (merror
      (intl:gettext "properties: argument must be a symbol or a string.")))
  (let ((u (properties x)) (v (or (safe-get x 'noun) (safe-get x 'verb))))
    (if v (nconc u (cdr (properties v))) u)))

(defun properties (x)
  (if (stringp x)
    ; AT THIS POINT WE MIGHT WANT TO TRY TO TEST ALL CHARS IN STRING ...
    (if (and (> (length x) 0) (member (char x 0) *alphabet*))
      '((mlist) $alphabetic)
      '((mlist)))
    (do ((y (symbol-plist x) (cddr y))
	 (l (cons '(mlist simp) (and (boundp x)
  				   (if (optionp x) (ncons "system value")
  				       (ncons '$value)))))
	 (prop))
	((null y)
	 (if (member x (cdr $features) :test #'eq) (nconc l (ncons '$feature)))
	 (if (member x (cdr $contexts) :test #'eq) (nconc l (ncons '$context)))
	 (if (member x (cdr $activecontexts) :test #'eq)
  	   (nconc l (ncons '$activecontext)))
	 (cond  ((null (symbol-plist x))
  	       (if (fboundp x) (nconc l (list "system function")))))
	 l)

      ;; TOP-LEVEL PROPERTIES
      (cond ((setq prop 
                   (assoc (car y)
                          `((bindtest . $bindtest)
                            (sp2 . $deftaylor)
                            (sp2subs . $deftaylor)
                            (assign . "assign property")
                            (nonarray . $nonarray)
                            (grad . $gradef)
                            (integral . $integral)
                            (distribute_over . "distributes over bags")
                            (simplim%function . "limit function")
                            (conjugate-function . "conjugate function")
                            (commutes-with-conjugate . "mirror symmetry")
  			    (risplit-function . "complex characteristic")
                            (noun . $noun)
                            (evfun . $evfun)
                            (evflag . $evflag)
                            (op . $operator)) :test #'eq))
  	   (nconc l (ncons (cdr prop))))
  	  ((setq prop (member (car y) opers :test #'eq)) 
  	   (nconc l (list (car prop))))
  	  ((and (eq (car y) 'operators) (not (or (eq (cadr y) 'simpargs1) (eq (cadr y) nil))))
  	   (nconc l (list '$rule)))
  	  ((and (member (car y) '(fexpr fsubr mfexpr*s mfexpr*) :test #'eq)
  		(nconc l (ncons "special evaluation form"))
  		nil))
  	  ((and (or (get (car y) 'mfexpr*) (fboundp x))
  	        ;; Do not add more than one entry to the list.
  	        (not (member '$transfun l))
  	        (not (member '$rule l))
  		(not (member "system function" l :test #'equal)))
  	   (nconc l
  		  (list (cond ((get x 'translated) '$transfun)
  			      ((mgetl x '($rule ruleof)) '$rule)
  			      (t "system function")))))
  	  ((and (eq (car y) 'autoload) 
  	        (not (member "system function" l :test #'equal)))
  	   (nconc l (ncons (if (member x (cdr $props) :test #'eq)
  			       "user autoload function"
  			       "system function"))))
  	  ((and (eq (car y) 'reversealias) 
  	        (member (car y) (cdr $aliases) :test #'eq))
  	   (nconc l (ncons '$alias)))
  	  ((eq (car y) 'data)
  	   (nconc l (cons "database info" (cdr ($facts x)))))
  	  ((eq (car y) 'mprops)
  	   ;; PROPS PROPERTIES
  	   (do ((y
  		 (cdadr y)
  		 (cddr y)))
  	       ((null y))
  	     (cond ((setq prop (assoc (car y)
  				     `((mexpr . $function)
  				       (mmacro . $macro)
  				       (hashar . "hashed array")
  				       (aexpr . "array function")
  				       (atvalues . $atvalue)
  				       ($atomgrad . $atomgrad)
  				       ($numer . $numer)
  				       (depends . $dependency)
  				       ($nonscalar . $nonscalar)
  				       ($scalar . $scalar)
  				       (matchdeclare . $matchdeclare)
  				       (mode . $modedeclare)) :test #'eq))
  		    (nconc l (list (cdr prop))))
  		   ((eq (car y) 'array)
  		    (nconc l
  			   (list (cond ((get x 'array) "complete array")
  				       (t "declared array")))))
  		   ((and (eq (car y) '$props) (cdadr y))
  		    (nconc l
  			   (do ((y (cdadr y) (cddr y))
  				(l (list '(mlist) "user properties")))
  			       ((null y) (list l))
  			     (nconc l (list (car y)))))))))))))

(defmspec $propvars (x)
  (setq x (fexprcheck x))
  (do ((iteml (cdr $props) (cdr iteml)) (propvars (ncons '(mlist))))
      ((null iteml) propvars)
    (and (among x (meval (list '($properties) (car iteml))))
	 (nconc propvars (ncons (car iteml))))))

(defmspec $printprops (r) (setq r (cdr r))
	  (if (null (cdr r)) (merror (intl:gettext "printprops: requires two arguments.")))
	  (let ((s (cadr r)))
	    (setq r (car r))
	    (setq r (cond ((atom r)
			   (cond ((eq r '$all)
				  (cond ((eq s '$gradef) (mapcar 'caar (cdr $gradefs)))
					(t (cdr (meval (list '($propvars) s))))))
				 (t (ncons r))))
			  (t (cdr r))))
	    (cond ((eq s '$atvalue) (dispatvalues r))
		  ((eq s '$atomgrad) (dispatomgrads r))
		  ((eq s '$gradef) (dispgradefs r))
		  ((eq s '$matchdeclare) (dispmatchdeclares r))
		  (t (merror (intl:gettext "printprops: unknown property ~:M") s)))))

(defun dispatvalues (l)
  (do ((l l (cdr l)))
      ((null l))
    (do ((ll (mget (car l) 'atvalues) (cdr ll)))
	((null ll))
      (mtell-open "~M~%"
		  (list '(mlabel) nil
			(list '(mequal)
			      (atdecode (car l) (caar ll) (cadar ll)) (caddar ll))))))
  '$done)

(defun atdecode (fun dl vl)
  (setq vl (copy-list vl))
  (atvarschk vl)
  (let ((eqs nil) (nvarl nil))
    (cond ((not (member nil (mapcar #'(lambda (x) (signp e x)) dl) :test #'eq))
	   (do ((vl vl (cdr vl)) (varl atvars (cdr varl)))
	       ((null vl))
	     (and (eq (car vl) munbound) (rplaca vl (car varl))))
	   (cons (list fun) vl))
	  (t (setq fun (cons (list fun)
			     (do ((n (length vl) (1- n))
				  (varl atvars (cdr varl))
				  (l nil (cons (car varl) l)))
				 ((zerop n) (nreverse l)))))
	     (do ((vl vl (cdr vl)) (varl atvars (cdr varl)))
		 ((null vl))
	       (and (not (eq (car vl) munbound))
		    (setq eqs (cons (list '(mequal) (car varl) (car vl)) eqs))))
	     (setq eqs (cons '(mlist) (nreverse eqs)))
	     (do ((varl atvars (cdr varl)) (dl dl (cdr dl)))
		 ((null dl) (setq nvarl (nreverse nvarl)))
	       (and (not (zerop (car dl)))
		    (setq nvarl (cons (car dl) (cons (car varl) nvarl)))))
	     (list '(%at) (cons '(%derivative) (cons fun nvarl)) eqs)))))

(defun dispatomgrads (l)
  (do ((i l (cdr i)))
      ((null i))
    (do ((j (mget (car i) '$atomgrad) (cdr j)))
	((null j))
      (mtell-open "~M~%"
		  (list '(mlabel) nil
			(list '(mequal)
			      (list '(%derivative) (car i) (caar j) 1) (cdar j))))))
  '$done)

(defun dispgradefs (l)
  (do ((i l (cdr i)))
      ((null i))
    (setq l (get (car i) 'grad))
    (do ((j (car l) (cdr j))
	 (k (cdr l) (cdr k))
	 (thing (cons (ncons (car i)) (car l))))
	((or (null k) (null j)))
      (mtell-open "~M~%"
		  (list '(mlabel)
			nil (list '(mequal) (list '(%derivative) thing (car j) 1.) (car k))))))
  '$done)

(defun dispmatchdeclares (l)
  (do ((i l (cdr i))
       (ret))
      ((null i) (cons '(mlist) (reverse ret)))
    (setq l (car (mget (car i) 'matchdeclare)))
    (setq ret (cons (append (cond ((atom l) (ncons (ncons l))) ((eq (caar l) 'lambda) (list '(mqapply) l))  (t l))
			    (ncons (car i)))
		    ret))))

(declare-top (special $programmode *roots *failures varlist genvar $ratfac))

(defmfun $changevar (expr trans nvar ovar)
  (let ($ratfac)
    (cond ((or (atom expr) (eq (caar expr) 'rat) (eq (caar expr) 'mrat))
	   expr)
	  ((atom trans)
	   (merror (intl:gettext "changevar: second argument must not be an atom; found: ~M") trans))
	  ((null (atom nvar))
	   (merror (intl:gettext "changevar: third argument must be an atom; found: ~M") nvar))
	  ((null (atom ovar))
	   (merror (intl:gettext "changevar: fourth argument must be an atom; found: ~M") ovar)))
    (changevar expr trans nvar ovar)))

(defun solvable (l var &optional (errswitch nil))
  (let (*roots *failures)
    (solve l var 1)
    (cond (*roots
	   ;; We arbitrarily pick the first root.  Should we be more careful?
	   ($rhs (car *roots)))
	  (errswitch (merror (intl:gettext "changevar: failed to solve for ~M in ~M") var l))
	  (t nil))))

(defun changevar (expr trans nvar ovar)
  (cond ((atom expr) expr)
	((or (not (member (caar expr) '(%integrate %sum %product) :test #'eq))
	     (not (alike1 (caddr expr) ovar)))
	 (recur-apply (lambda (e) (changevar e trans nvar ovar)) expr))
	(t
	 ;; TRANS is the expression that relates old var and new var
	 ;; and is of the form f(ovar, nvar) = 0. Using TRANS, try to
	 ;; solve for ovar so that ovar = tfun(nvar), if possible.
	 (let* ((tfun (solvable (setq trans (meqhk trans)) ovar))
		(deriv
		 ;; Compute diff(tfun, nvar) = dovar/dnvar if tfun is
		 ;; available.  Otherwise, use implicit
		 ;; differentiation.
		 (if tfun
		     (sdiff tfun nvar)
		     (neg (div (sdiff trans nvar) ;IMPLICIT DIFF.
			       (sdiff trans ovar)))))
		(sum-product-p (member (caar expr) '(%sum %product) :test #'eq)))

	   #+nil
	   (progn
	     (mformat t "tfun = ~M~%" tfun)
	     (mformat t "deriv = ~M~%" deriv))

	   ;; For sums and products, we want deriv to be +/-1 because
	   ;; I think that means that integers will map into integers
	   ;; (roughly), so that we don't need to express the
	   ;; summation index or limits in some special way to account
	   ;; for it.
	   (when (and (member (caar expr) '(%sum %product) :test #'eq)
		      (not (or (equal deriv 1)
			       (equal deriv -1))))
	     (merror (intl:gettext "changevar: illegal change in summation or product")))

	   (let ((nfun ($radcan	;NIL IF KERNSUBST FAILS
			      (if tfun
				  (mul (maxima-substitute tfun ovar (cadr expr))
				       ;; Don't multiply by deriv
				       ;; for sums/products because
				       ;; reversing the order of
				       ;; limits doesn't change the
				       ;; sign of the result.
				       (if sum-product-p 1 deriv))
				  (kernsubst ($ratsimp (mul (cadr expr)
							    deriv))
					     trans ovar)))))
	     (cond
	       (nfun
		;; nfun is basically the result of subtituting ovar
		;; with tfun in the integratand (summand).
		(cond	
		  ((cdddr expr)
		   ;; Handle definite integral, summation, or product.
		   ;; invfun expresses nvar in terms of ovar so that
		   ;; we can compute the new lower and upper limits of
		   ;; the integral (sum).
		   (let* ((invfun (solvable trans nvar t))
			  (lo-limit ($limit invfun ovar (cadddr expr) '$plus))
			  (hi-limit ($limit invfun
					    ovar
					    (car (cddddr expr))
					    '$minus)))
		     ;; If this is a sum or product and deriv = -1, we
		     ;; want to reverse the low and high limits.
		     (when (and sum-product-p (equal deriv -1))
		       (rotatef lo-limit hi-limit))

		     ;; Construct the new result.
		     (list (ncons (caar expr))
			   nfun
			   nvar
			   lo-limit
			   hi-limit)))
		  (t
		   ;; Indefinite integral
		   (list '(%integrate) nfun nvar))))
	       (t expr)))))))

(defun kernsubst (expr form ovar)
  (let (varlist genvar nvarlist)
    (newvar expr)
    (setq nvarlist (mapcar #'(lambda (x) (if (freeof ovar x) x
					     (solvable form x)))
			   varlist))
    (if (member nil nvarlist :test #'eq) nil
	(prog2 (setq expr (ratrep* expr)
		     varlist nvarlist)
	    (rdis (cdr expr))))))

(declare-top (special $listconstvars facfun))

(defmfun $factorsum (e)
  (factorsum0 e '$factor))

(defmfun $gfactorsum (e)
  (factorsum0 e '$gfactor))

(defun factorsum0 (e facfun)
  (cond ((mplusp (setq e (funcall facfun e)))
	 (factorsum1 (cdr e)))
	(t (factorsum2 e))))

(defun factorsum1 (e)
  (prog (f lv llv lex cl lt c)
   loop (setq f (car e))
   (setq lv (cdr ($showratvars f)))
   (cond ((null lv) (setq cl (cons f cl)) (go skip)))
   (do ((q llv (cdr q)) (r lex (cdr r)))
       ((null q))
     (cond ((intersect (car q) lv)
	    (rplaca q (union* (car q) lv))
	    (rplaca r (cons f (car r)))
	    (return (setq lv nil)))))
   (or lv (go skip))
   (setq llv (cons lv llv) lex (cons (ncons f) lex))
   skip (and (setq e (cdr e)) (go loop))
   (or cl (go skip2))
   (do ((q llv (cdr q)) (r lex (cdr r)))
       ((null q))
     (cond ((and (null (cdar q)) (cdar r))
	    (rplaca r (nconc cl (car r)))
	    (return (setq cl nil)))))
   skip2 (setq llv nil lv nil)
   (do ((r lex (cdr r)))
       ((null r))
     (cond ((cdar r)
	    (setq llv
		  (cons (factorsum2 (funcall facfun (cons '(mplus) (car r))))
			llv)))
	   ((or (not (mtimesp (setq f (caar r))))
		(not (mnump (setq c (cadr f)))))
	    (setq llv (cons f llv)))
	   (t (do ((q lt (cdr q)) (s lv (cdr s)))
		  ((null q))
		(cond ((alike1 (car s) c)
		       (rplaca q (cons (dcon f) (car q)))
		       (return (setq f nil)))))
	      (and f
		   (setq lv (cons c lv)
			 lt (cons (ncons (dcon f)) lt))))))
   (setq lex
	 (mapcar #'(lambda (s q)
		     (simptimes (list '(mtimes) s
				      (cond ((cdr q)
					     (cons '(mplus) q))
					    (t (car q))))
				1 nil))
		 lv lt))
   (return (simplus (cons '(mplus) (nconc cl lex llv)) 1 nil))))

(defun dcon (mt)
  (cond ((cdddr mt) (cons (car mt) (cddr mt))) (t (caddr mt))))

(defun factorsum2 (e)
  (cond ((not (mtimesp e)) e)
	(t (cons '(mtimes)
		 (mapcar #'(lambda (f)
			     (cond ((mplusp f)
				    (factorsum1 (cdr f)))
				   (t f)))
			 (cdr e))))))

(declare-top (special $combineflag))

(defmvar $combineflag t)

(defmfun $combine (e)
  (cond ((or (atom e) (eq (caar e) 'rat)) e)
	((eq (caar e) 'mplus) (combine (cdr e)))
	(t (recur-apply #'$combine e))))

(defun combine (e)
  (prog (term r ld sw nnu d ln xl)
   again(setq term (car e) e (cdr e))
   (when (or (not (or (ratnump term) (mtimesp term) (mexptp term)))
	     (equal (setq d ($denom term)) 1))
     (setq r (cons term r))
     (go end))
   (setq nnu ($num term))
   (and $combineflag (integerp d) (setq xl (cons term xl)) (go end))
   (do ((q ld (cdr q)) (p ln (cdr p)))
       ((null q))
     (cond ((alike1 (car q) d)
	    (rplaca p (cons nnu (car p)))
	    (return (setq sw t)))))
   (and sw (go skip))
   (setq ld (cons d ld) ln (cons (ncons nnu) ln))
   skip (setq sw nil)
   end  (and e (go again))
   (and xl (setq xl (cond ((cdr xl) ($xthru (addn xl t)))
			  (t (car xl)))))
   (mapc
    #'(lambda (nu de)
	(setq r (cons (mul2 (addn nu nil) (power* de -1)) r)))
    ln ld)
   (return (addn (if xl (cons xl r) r) nil))))

(defmfun $factorout (e &rest vl)
  (prog (el fl cl l f x)
     (when (null vl)
       (merror (intl:gettext "factorout: at least two arguments required.")))
     (unless (mplusp e)
       (return e))
     (or (null vl) (mplusp e) (return e))
     (setq e (cdr e))
     loop	(setq f (car e) e (cdr e))
     (unless (mtimesp f)
       (setq f (list '(mtimes) 1 f)))
     (setq fl nil cl nil)
     (do ((i (cdr f) (cdr i)))
	 ((null i))
       (if (and (not (numberp (car i)))
		(apply '$freeof (append vl (ncons (car i)))))
	   (setq fl (cons (car i) fl))
	   (setq cl (cons (car i) cl))))
     (when (null fl)
       (push f el)
       (go end))
     (setq fl (if (cdr fl)
		  (simptimes (cons '(mtimes) fl) 1 nil)
		  (car fl)))
     (setq cl (cond ((null cl) 1)
		    ((cdr cl) (simptimes (cons '(mtimes) cl) 1 t))
		    (t (car cl))))
     (setq x t)
     (do ((i l (cdr i)))
	 ((null i))
       (when (alike1 (caar i) fl)
	 (rplacd (car i) (cons cl (cdar i)))
	 (setq i nil x nil)))
     (when x
       (push (list fl cl) l))
     end  (when e (go loop))
     (do ((i l (cdr i)))
	 ((null i))
       (push (simptimes (list '(mtimes) (caar i)
			      ($factorsum (simplus (cons '(mplus) (cdar i)) 1 nil))) 1 nil) el))
     (return (addn el nil))))

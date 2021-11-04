;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     The data in this file contains enhancments.                    ;;;;;
;;;                                                                    ;;;;;
;;;  Copyright (c) 1984,1987 by William Schelter,University of Texas   ;;;;;
;;;     All rights reserved                                            ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(in-package :maxima)

(defvar $manual_demo "manual.demo")

(defmspec $example (l)
  (declare (special *need-prompt*))
  (let ((example (second l)))
    (when (symbolp example)
      ;; Coerce a symbol to be a string.
      ;; Remove the first character if it is a dollar sign.
      (setq example (coerce (exploden (stripdollar example)) 'string)))
    (unless (stringp example)
      (merror 
        (intl:gettext "example: argument must be a symbol or a string; found: ~M") example))
    ;; Downcase the string. $example is not case sensitive.
    (setq example (string-downcase example))
    (with-open-file (st ($file_search1 $manual_demo '((mlist) $file_search_demo)))
      (prog (tem all c-tag d-tag)
       again
       (setq tem (read-char st nil))
       (unless tem (go notfound))
       (unless (eql tem #\&) (go again))
       (setq tem (read-char st nil))
       (unless (eql tem #\&) (go again))
       ;; so we are just after having read &&

       (setq tem (read st nil nil))
       (unless tem (go notfound))
       ;; Coerce the topic in tem to be a string.
       (setq tem (coerce (exploden tem) 'string))
       (cond ((string= tem example)
	      (go doit))
	     (t (push tem all)
		(go again)))
       ;; at this stage we read maxima forms and print and eval
       ;; until a peek sees '&' as the first character of next expression,
       ;; but at first skip over whitespaces.
       doit       
       (when (member (setq tem (peek-char nil st nil)) 
                     '(#\tab #\space #\newline #\linefeed #\return #\page))
         ;; Found whitespace. Read char and look for next char.
         ;; The && label can be positioned anywhere before the next topic.
         (setq tem (read-char st nil))
         (go doit))
       (cond ((or (null tem) (eql tem #\&))
	      (setf *need-prompt* t)
	      (return '$done)))
       (setq tem (dbm-read st nil nil))
       (incf $linenum)
       (setq c-tag (makelabel $inchar))
       (unless $nolabels (setf (symbol-value c-tag) (nth 2 tem)))
       (let ($display2d)
	 (displa `((mlabel) ,c-tag ,(nth 2 tem))))
       (setq $% (meval* (nth 2 tem)))
       (setq d-tag (makelabel $outchar))
       (unless $nolabels (setf (symbol-value d-tag) $%))
       (when (eq (caar tem) 'displayinput)
	 (displa `((mlabel) ,d-tag ,$%)))
       (go doit)

       notfound
       (setf *need-prompt* t)
       (if (= (length l) 1)
         (return `((mlist) ,@(nreverse all)))
         (progn
           (mtell (intl:gettext "example: ~M not found. 'example();' returns the list of known examples.~%") example)
           (return '$done)))))))

(defun mread-noprompt (&rest read-args)
  (let ((*mread-prompt* "") (*prompt-on-read-hang*))
    (declare (special *mread-prompt* *prompt-on-read-hang*))
    (unless read-args (setq read-args (list #+(or sbcl cmu) *standard-input*
                                            #-(or sbcl cmu) *query-io*)))
    (caddr (apply #'mread read-args))))

;; Some list creation utilities.

(defmspec $create_list (l)
  (cons '(mlist) (apply #'create-list1 (cadr l) (cddr l))))

(defun create-list1 (form &rest l &aux lis var1 top)
  (cond ((null l) (list (meval* form)))
	(t
	 (setq var1 (first l)
	       lis (second l)
	       l (cddr l))
	 (unless (symbolp var1) (merror (intl:gettext "create_list: expected a symbol; found: ~A") var1))
 	 (setq lis (meval* lis))
	 (mbinding ((list var1))
	   (cond ((and (numberp lis)
		       (progn
			 (setq top (car l) l (cdr l))
			 (setq top (meval* top))
			 (numberp top)))
		  (loop for i from lis to top
		     do (mset var1 i)
		     append
		     (apply #'create-list1 form l)))
		 (($listp lis)
		  (loop for v in (cdr lis)
		     do (mset var1 v)
		     append
		     (apply #'create-list1 form l)))
		 (t (merror (intl:gettext "create_list: unexpected arguments."))))))))

;; The documentation is now in INFO format and can be printed using
;; tex, or viewed using info or gnu emacs or using a web browser.  All
;; versions of maxima have a builtin info retrieval mechanism.

(defmspec $describe (x)
  (let ((topic ($sconcat (cadr x)))
	(exact-p (or (null (caddr x)) (eq (caddr x) '$exact))))
    (if exact-p
	(cl-info::info-exact topic)
	(cl-info::info-inexact topic))))

(defmspec $apropos (s)
  (setq s (car (margs s)))
  (cond ((or (stringp s)
	     (and (symbolp s) (setq s (string (stripdollar s)))))
         ;; A list of all Maxima names which contain the string S.
         (let ((acc (apropos-list s :maxima)))
           ;; Filter the names which are Maxima User symbols starting
           ;; with % or $ and remove duplicates.
           (remove-duplicates
            (cons '(mlist)
                   (delete-if-not
		    #'(lambda (x)
                        (cond ((eq x '||) nil)
			      ((char= (get-first-char x) #\$) x)
                              ;; Change to a verb, when present.
                              ((char= (get-first-char x) #\%) (or (get x 'noun) x))
                              (t nil)))
                    acc)) :test #'eq)))
        (t
         (merror
          (intl:gettext "apropos: argument must be a string or symbol; found: ~M") s))))

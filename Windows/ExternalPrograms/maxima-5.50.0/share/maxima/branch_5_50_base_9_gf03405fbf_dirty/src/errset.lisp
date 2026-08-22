;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     The data in this file contains enhancements.                   ;;;;;
;;;                                                                    ;;;;;
;;;  Copyright (c) 1984,1987 by William Schelter,University of Texas   ;;;;;
;;;     All rights reserved                                            ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; Macros and conditions related to signaling Maxima errors.  These
;;; use to be in merror.lisp but were moved here to make the
;;; dependencies between defsystem modules simpler.

(in-package :maxima)

(define-condition maxima-$error (error)
  ((message :initform $error :reader the-$error))
  (:documentation "Muser error, to be signalled by MERROR, usually.")
  (:report (lambda (c stream)
	     (declare (ignore c))
	     (let ((*standard-output* stream))
	       ($errormsg)))))

(defvar *merror-signals-$error-p* nil
  "When T, MERROR will signal a MAXIMA-$ERROR condition.")

;; Sample:
;; (defun h (he)
;;   (merror "hi there ~:M and ~:M" he he))
;; This will signal a MAXIMA-$ERROR condition:
;; (with-$error (h '$you))

(defmacro with-$error (&body body)
  "Let MERROR signal a MAXIMA-$ERROR condition."
  `(let ((*merror-signals-$error-p* t))
     (declare (special *merror-signals-$error-p*))
     ,@body))

(defmacro rat-error-to-merror (&body body)
  (let ((result (gensym)) (error-args (gensym)) (error-p (gensym)))
    `(let ((,result) (,error-p t))
       (let ((,error-args
              (catch 'rat-err
                (setf ,result (progn ,@body))
                (setf ,error-p nil))))
         (when ,error-p
           (apply #'merror ,error-args)))
       ,result)))

; Evaluate form while catching throws to some specific tags (called
; "errcatch tags").  If no throw to an errcatch tag is caught, then
; the values from form are returned.  If a throw to an errcatch tag
; is caught, then a Maxima error is signaled.
;
; The errcatch tags are ERRORSW, MACSYMA-QUIT and RAT-ERR.
(defmacro with-errcatch-tag-$errors (form)
  (let ((block-name (gensym)))
    `(block ,block-name
       ; RAT-ERROR-TO-MERROR will catch any throws to RAT-ERR and
       ; call merror with a specific error message.
       (catch 'macsyma-quit
         (catch 'errorsw
           (rat-error-to-merror
             (return-from ,block-name ,form))))
       ; If we're here, then we don't know any information about the
       ; error, so just call MERROR with a vague error message.  This
       ; message will not be printed by MERROR, but it will be stored
       ; in Maxima's error variable.
       (with-$error
         (merror (intl:gettext "An error was caught by errcatch."))))))

; This is similar to the classic errset, but errcatch handles lisp and
; Maxima errors.
(defmacro errcatch (form)
  `(let ((errcatch (cons bindlist loclist))
         (*mdebug* nil))
     (declare (special errcatch))
     (handler-case (list (with-errcatch-tag-$errors ,form))
       (maxima-$error ()
         ; If this was signaled by MERROR, then it has already handled
         ; the setting of the error variable and the printing of any error
         ; messages (as applicable).
         ;
         ; If for some reason this wasn't signaled by MERROR, then it's the
         ; signaler's responsibility to handle error messages.
         ;
         ; Either way, we just need to clean up here.
         (errlfun1 errcatch)
         nil)
       (error (e)
         ; We store the error report message in the error variable and
         ; print the message if errormsg is true.  Then we clean up.
         (setq $error (list '(mlist simp) (princ-to-string e)))
         (when $errormsg
           ($errormsg))
         (errlfun1 errcatch)
         nil))))

;;; -*-  Mode: Lisp; Package: Maxima; Syntax: Common-Lisp; Base: 10 -*- ;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;     The data in this file contains enhancments.                    ;;;;;
;;;                                                                    ;;;;;
;;;  Copyright (c) 1984,1987 by William Schelter,University of Texas   ;;;;;
;;;     All rights reserved                                            ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; ** (c) Copyright 1981 Massachusetts Institute of Technology **
;;;
;;; SYSTEM: The ``New'' Macsyma System Stuff

(in-package :maxima)

(macsyma-module system)

(defmvar $showtime nil
  "When T, the computation time is printed with each output expression.")

;;; Standard Kinds of Input Prompts

(defmvar $prompt '_
  "Prompt symbol of the demo function, playback, and the Maxima break loop.")

;; A prefix and suffix that are wrapped around every prompt that Maxima
;; emits. This is designed for use with text-based interfaces that drive Maxima
;; through standard input and output and need to decorate prompts to make the
;; output easier to parse. There are some more notes in
;; doc/implementation/external-interface.txt.
(defvar *prompt-prefix* "")
(defvar *prompt-suffix* "")
(defvar *general-display-prefix* "")
(defvar $alt_format_prompt nil "If NIL, use DEFAULT-FORMAT-PROMPT to print input prompt; if a function, use it to print input prompt.")

(defun format-prompt (destination control-string &rest arguments)
  "If $ALT_FORMAT_PROMPT is NIL, use DEFAULT-FORMAT-PROMPT to print
prompt; otherwise MFUNCALL $ALT_FORMAT_PROMPT to print prompt."
  (funcall (if $alt_format_prompt #'alt-format-prompt #'default-format-prompt)
	   destination control-string arguments))

(defun alt-format-prompt (destination control-string arguments)
  "MFUNCALL $ALT_FORMAT_PROMPT with a heavy coating of error protection."
  (handler-bind ((error (lambda(msg) (setq $alt_format_prompt nil)
			       (format t (intl:gettext "Error in printing prompt; reverting to default.~%~a") msg)
			       (throw 'macsyma-quit 'maxima-error))))
    (with-$error (let ((prompt (mfuncall $alt_format_prompt destination control-string arguments)))
		   (if (stringp prompt) prompt (merror "alt_format_prompt returned an object of type ~a, needed a string." (type-of prompt)))))))

(defun default-format-prompt (destination control-string arguments)
  "Like AFORMAT, but add the prefix and suffix configured for a prompt. This
function deals correctly with the ~M control character, but only when
DESTINATION is an actual stream (rather than nil for a string)."
  (let ((*print-circle* nil))
    (if (null destination)
	;; return value string is important
	(concatenate 'string
		     *prompt-prefix*		     
		     (apply #'aformat destination
			    control-string
			    arguments)
		     *prompt-suffix*)
      (progn
	(format destination "~A~A~A"
		*prompt-prefix*		     
		(apply #'aformat nil
		       control-string
		       arguments)
		*prompt-suffix*)))))
    

(defvar $default_format_prompt (symbol-function 'default-format-prompt))

;;  "When time began" (or at least the start of version control history),
;;  the following comment was made at this point:
;;
;;     instead of using this STRIPDOLLAR hackery, the
;;     MREAD function should call MFORMAT to print the prompt,
;;     and take a format string and format arguments.
;;     Even easier and more general is for MREAD to take
;;     a FUNARG as the prompt. -gjc
;;
;;  I guess we're still failing miserably, but unfortunately MFORMAT/AFORMAT
;;  don't deal correctly with ~M plus a string output stream.
(defun main-prompt ()
  (declare (special *display-labels-p*))
  (if *display-labels-p*
      (format-prompt nil "(~A~A) "
                     (print-invert-case (stripdollar $inchar))
                     $linenum)
      ""))

(defun break-prompt ()
  (format-prompt nil "~A"
                 (print-invert-case (stripdollar $prompt))))

(defun toplevel-macsyma-eval (x)
  ;; Catch rat-err's here.
  ;;
  ;; The idea is that eventually there will be quite a few "maybe catch this"
  ;; errors, which will be raised and might well get eaten before they get as far
  ;; as here. However, we want to display them nicely like merror rather than
  ;; letting a lisp error percolate to the debugger and, as such, we catch them
  ;; here and replace them with an merror call.
  ;;
  ;; Other random errors get to the lisp debugger, which is normally set to print
  ;; them and continue, via *debugger-hook*.
  (rat-error-to-merror (meval* x)))

(defmvar $_ '$_ "last thing read in, corresponds to lisp +")
(defmvar $__ '$__ "thing read in which will be evaluated, corresponds to -")

(declare-top (special *mread-prompt*  $file_search_demo))

(defvar accumulated-time 0.0)

#+(or cmu scl)
(defun used-area (&optional unused)
  (declare (ignore unused))
  (ext:get-bytes-consed))

#+sbcl
(defun used-area (&optional unused)
  (declare (ignore unused))
  (sb-ext:get-bytes-consed))

#+openmcl
(defun used-area (&optional unused)
  (declare (ignore unused))
  (ccl::total-bytes-allocated))

#+clisp
(defun used-area (&optional unused)
  (declare (ignore unused))
  (multiple-value-bind (real1 real2 run1 run2 gc1 gc2 space1 space2 gccount)
      (sys::%%time)
    (declare (ignore real1 real2 run1 run2 gc1 gc2 gccount))
    (dpb space1 (byte 24 24) space2)))


#+allegro
(defun used-area (&optional unused)
  (declare (ignore unused))
  (declare (optimize (speed 3)))
  (let ((.oldspace (make-array 4 :element-type
			       #-64bit '(unsigned-byte 32)
			       #+64bit '(unsigned-byte 64))))
    (declare (type (simple-array #-64bit (unsigned-byte 32)
				 #+64bit (unsigned-byte 64) (*))
		   .oldspace))

    (multiple-value-bind (.olduser .oldsystem .oldgcu .oldgcs)
	(excl::get-internal-run-times)
      (declare (ignore .olduser .oldsystem .oldgcs))
      (sys::gsgc-totalloc .oldspace t)
      (list (aref .oldspace 0) (aref .oldspace 2) .oldgcu)))) ;; report just two kinds of space,
							      ;; cons-cells and other bytes,
							      ;; also report gc-user time

#+lispworks
(defun used-area (&optional unused)
  (declare (ignore unused))
  (getf (system:room-values) :total-allocated))

#-(or cmu scl sbcl clisp allegro openmcl lispworks)
(defun used-area (&optional unused)
  (declare (ignore unused))
  0)

(defun continue (&optional (input-stream *standard-input*)
		 batch-or-demo-flag)
  (declare (special *socket-connection*))
  (if (eql batch-or-demo-flag :demo)
      (format t
        (intl:gettext
          "~%At the '~A' prompt, type ';' and <enter> to get next demonstration.~&")
        (print-invert-case (stripdollar $prompt))))
  (catch 'abort-demo
    (do ((r)
	 (time-before)
	 (time-after)
	 (time-used)
	 (eof (list nil))
	 (etime-before)
	 (etime-after)
	 (area-before)
	 (area-after)
	 (etime-used)
	 (c-tag)
	 (d-tag))
	(nil)
      (catch 'return-from-debugger
	(when (or (not (checklabel $inchar))
		  (not (checklabel $outchar)))
	  (incf $linenum))
	(setq c-tag (makelabel $inchar))
	(let ((*mread-prompt* (if batch-or-demo-flag nil (main-prompt)))
	      (eof-count 0))
	  (tagbody
	   top
	     (setq r (dbm-read input-stream nil eof))
	     ;; This is something of a hack. If we are running in a server mode
	     ;; (which we determine by checking *socket-connection*) and we get
	     ;; an eof on an input-stream that is not *standard-input*, switch
	     ;; the input stream to *standard-input*.
	     ;; There should probably be a better scheme for server mode.
	     ;; jfa 10/09/2002.
	     (if (and
		  (eq r eof)
		  (not (eq input-stream *standard-input*))
		  (boundp '*socket-connection*))
		 (progn
		   (setq input-stream *standard-input*)
		   (if batch-or-demo-flag
		       (return '$done)
		       (progn
			 (setq *mread-prompt* nil)
			 (setq r (dbm-read input-stream nil eof))))))

	     (cond ((and (eq r eof) (boundp '*socket-connection*)
			 (eq input-stream *socket-connection*))
		    (cond ((>=  (setq eof-count (+ 1 eof-count)) 10)
			   (print "exiting on eof")
			   ($quit))
			  (t (go top)))))
	     (cond ((and (consp r) (keywordp (car r)))
		    (break-call (car r) (cdr r) 'break-command)
		    (go top)))))
	(format t "~a" *general-display-prefix*)
	(if (eq r eof) (return '$done))
	(setq $__ (caddr r))
	(unless $nolabels (setf (symbol-value c-tag) $__))
	(cond (batch-or-demo-flag
	  (let (($display2d nil))
	    (displa `((mlabel) ,c-tag , $__)))))
	(setq time-before (get-internal-run-time)
	      etime-before (get-internal-real-time))
	(setq area-before (used-area))
	(setq $% (toplevel-macsyma-eval $__))
	(setq etime-after (get-internal-real-time)
	      time-after (get-internal-run-time))
	(setq area-after (used-area))
	(setq time-used (quotient
			 (float (- time-after time-before))
			 internal-time-units-per-second)
	      etime-used (quotient
			  (float (- etime-after etime-before))
			  internal-time-units-per-second))
	(incf accumulated-time time-used)
	(setq d-tag (makelabel $outchar))
	(unless $nolabels (setf (symbol-value d-tag) $%))
	(setq $_ $__)
	(when $showtime	;; we don't distinguish showtime:all?? /RJF
	  (format t (intl:gettext "Evaluation took ~,4F seconds (~,4F elapsed)")
		  time-used etime-used )
	  #+(or gcl ecl)
	  (format t "~%")
	  #+(or cmu scl sbcl clisp openmcl)
	  (let ((total-bytes (- area-after area-before)))
	    (cond ((> total-bytes (* 1024 1024))
		   (format t (intl:gettext " using ~,3F MB.~%")
			   (/ total-bytes (* 1024.0 1024.0))))
		  ((> total-bytes 1024)
		   (format t (intl:gettext " using ~,3F KB.~%") (/ total-bytes 1024.0)))
		  (t
		   (format t (intl:gettext " using ~:D bytes.~%") total-bytes))))

	  #+allegro
	  (let ((conses (- (car area-after) (car area-before)))
		(other (- (cadr area-after) (cadr area-before)))
		(gctime (- (caddr area-after) (caddr area-before))))
	    (if (= 0 gctime) nil (format t (intl:gettext " including GC time ~s s,") (* 0.001 gctime)))
	    (format t (intl:gettext " using ~s cons-cells and ~s other bytes.~%") conses other)))
	(unless $nolabels
          (putprop '$% (cons time-used 0) 'time)
	  (putprop d-tag (cons time-used  0) 'time))
	(if (eq (caar r) 'displayinput)
	    (displa `((mlabel) ,d-tag ,$%)))
	(when (eq batch-or-demo-flag ':demo)
          (princ (break-prompt))
          (force-output)
	  (let (quitting)
	    (do ((char)) (nil)
	      ;;those are common lisp characters you're reading here
	      (case (setq char (read-char *terminal-io*))
                ((#\page)
                 (fresh-line)
                 (princ (break-prompt))
                 (force-output))
                ((#\?)
                 (format t
                   (intl:gettext
                     "  Pausing. Type a ';' and <enter> to continue demo.~%")))
		((#\space #\; #\n #\e #\x #\t))
		((#\newline )
		 (if quitting (throw 'abort-demo nil) (return nil)))
		(t (setq quitting t))))))
	;; This is sort of a kludge -- eat newlines and blanks so that
	;; they don't echo
	(and batch-or-demo-flag
	     (do ((char)) (())
	       (setq char (read-char input-stream nil nil))
	       (when (null char)
		 (throw 'macsyma-quit nil))
	       (unless (member char '(#\space #\newline #\return #\tab) :test #'equal)
		 (unread-char char input-stream)
		 (return nil))))))))

(defun $break (&rest arg-list)
  (prog1 (apply #'$print arg-list)
    (mbreak-loop)))

(defun mbreak-loop ()
  (let ((*standard-input* *debug-io*)
	(*standard-output* *debug-io*))
    (catch 'break-exit
      (format t (intl:gettext "~%Entering a Maxima break point. Type 'exit;' to resume."))
      (do ((r)) (nil)
	(fresh-line)
	(setq r (caddr (let ((*mread-prompt* (break-prompt)))
			 (mread *standard-input*))))
	(case r
	  (($exit) (throw 'break-exit t))
	  (t (errset (displa (meval r)) t)))))))

(defun merrbreak (&optional arg)
  (format *debug-io* "~%Merrbreak:~A" arg)
  (mbreak-loop))

(defun retrieve (msg flag &aux (print? nil))
  (declare (special msg flag print?))
  (or (eq flag 'noprint) (setq print? t))
  (cond ((not print?)
	 (setq print? t)
         (format-prompt t ""))
	((null msg)
         (format-prompt t ""))
	((atom msg)
         (format-prompt t "~A" msg)
	 (terpri))
	((eq flag t)
         (format-prompt t "~{~A~}" (cdr msg))
	 (mterpri))
	(t
         (format-prompt t "~M" msg)
	 (mterpri)))
  (let ((res (mread-noprompt *query-io* nil)))
    (princ *general-display-prefix*)
    res))

(defmfun $read (&rest l)
  (meval (apply #'$readonly l)))

(defmfun $readonly (&rest l)
  (let ((*mread-prompt*
	 (if l
	     (string-right-trim '(#\n)
				(with-output-to-string (*standard-output*) (apply #'$print l)))
	     "")))
    (setf *mread-prompt* (format-prompt nil "~A" *mread-prompt*))
    (third (mread *query-io*))))

;; FUNCTION BATCH APPARENTLY NEVER CALLED. OMIT FROM GETTEXT SWEEP AND DELETE IT EVENTUALLY
(defun batch (filename &optional demo-p
	      &aux (orig filename) list
	      file-obj (accumulated-time 0.0) (abortp t))
  (setq list (if demo-p '$file_search_demo '$file_search_maxima))
  (setq filename ($file_search filename (symbol-value list)))
  (or filename (merror "Could not find ~M in ~M: ~M"
		       orig list (symbol-value list)))

  (unwind-protect
       (progn (batch-internal (setq file-obj (open filename)) demo-p)
	      (setq abortp nil)
	      (when $showtime
		(format t "~&Batch spent ~,4F seconds in evaluation.~%"
			accumulated-time)))
    (if file-obj (close file-obj))
    (when abortp (format t "~&(Batch of ~A aborted.)~%" filename))))


(defun batch-internal (fileobj demo-p)
  (continue (make-echo-stream fileobj *standard-output*)
	    (if demo-p ':demo ':batch)))

(defun $demo (&rest arg-list)
  (let ((tem ($file_search (car arg-list) $file_search_demo)))
    (or tem (merror (intl:gettext "demo: could not find ~M in ~M.")
		    (car arg-list) '$file_search_demo))
    ($batch tem	'$demo)))

(defmfun $bug_report ()
  (format t (intl:gettext "~%Please report bugs to:~%"))
  (format t "    http://sourceforge.net/p/maxima/bugs~%")
  (format t (intl:gettext "To report a bug, you must have a Sourceforge account.~%"))
  (format t (intl:gettext "Please include the following information with your bug report:~%"))
  (format t "-------------------------------------------------------------~%")
  ; Display the 2D-formatted build information
  (let (($display2d t))
    (displa ($build_info)))
  (format t "-------------------------------------------------------------~%")
  (format t (intl:gettext "The above information is also reported by the function 'build_info()'.~%~%"))
  "")

;; Declare a build_info structure, then remove it from the list of user-defined structures.
(defstruct1 '((%build_info) $version $timestamp $host $lisp_name $lisp_version))
(let nil (declare (special $structures))
  (setq $structures (cons '(mlist) (remove-if #'(lambda (x) (eq (caar x) '%build_info)) (cdr $structures)))))

(defvar *maxima-build-info* nil)

(defmfun $build_info ()
  (or
    *maxima-build-info*
    (setq
      *maxima-build-info*
      (let
        ((year (sixth cl-user:*maxima-build-time*))
         (month (fifth cl-user:*maxima-build-time*))
         (day (fourth cl-user:*maxima-build-time*))
         (hour (third cl-user:*maxima-build-time*))
         (minute (second cl-user:*maxima-build-time*))
         (seconds (first cl-user:*maxima-build-time*)))
        (mfuncall
          '$new
          `((%build_info)
            ,*autoconf-version*
            ,(format nil "~4,'0d-~2,'0d-~2,'0d ~2,'0d:~2,'0d:~2,'0d" year month day hour minute seconds)
            ,*autoconf-host*
            ,#+sbcl (ensure-readably-printable-string (lisp-implementation-type)) #-sbcl (lisp-implementation-type)
            ,#+sbcl (ensure-readably-printable-string (lisp-implementation-version)) #-sbcl (lisp-implementation-version)))))))

;; SBCL base strings aren't readably printable.
;; Attempt a work-around. Yes, this is terribly ugly.
#+sbcl (defun ensure-readably-printable-string (x)
  (coerce x `(simple-array character (,(length x)))))

(defun dimension-build-info (form result)
  (declare (special bkptht bkptdp lines break))
  ;; Usually the result of (MFUNCALL '$@ ...) is a string,
  ;; but ensure that output makes sense even if it is not.
  (let
    ((version-string (format nil (intl:gettext "Maxima version: ~a")
       (coerce (mstring (mfuncall '$@ form '$version)) 'string)))
     (timestamp-string (format nil (intl:gettext "Maxima build date: ~a")
       (coerce (mstring (mfuncall '$@ form '$timestamp)) 'string)))
     (host-string (format nil (intl:gettext "Host type: ~a")
       (coerce (mstring (mfuncall '$@ form '$host)) 'string)))
     (lisp-name-string (format nil (intl:gettext "Lisp implementation type: ~a")
       (coerce (mstring (mfuncall '$@ form '$lisp_name)) 'string)))
     (lisp-version-string (format nil (intl:gettext "Lisp implementation version: ~a")
       (coerce (mstring (mfuncall '$@ form '$lisp_version)) 'string)))
     (bkptht 1)
     (bkptdp 1)
     (lines 0)
     (break 0))
    (forcebreak result 0)
    (forcebreak (reverse (coerce version-string 'list)) 0)
    (forcebreak (reverse (coerce timestamp-string 'list)) 0)
    (forcebreak (reverse (coerce host-string 'list)) 0)
    (forcebreak (reverse (coerce lisp-name-string 'list)) 0)
    (forcebreak (reverse (coerce lisp-version-string 'list)) 0))
  nil)

(setf (get '%build_info 'dimension) 'dimension-build-info)

(defvar *maxima-started* nil)

(defvar *maxima-prolog* "")
(defvar *maxima-epilog* "")

(declare-top (special *maxima-initmac* *maxima-initlisp*))

(defvar *maxima-quiet* nil)

(defun macsyma-top-level (&optional (input-stream *standard-input*) batch-flag)
  (let ((*package* (find-package :maxima)))
    (if *maxima-started*
	(format t (intl:gettext "Maxima restarted.~%"))
	(progn
	  (if (not *maxima-quiet*) (maxima-banner))
	  (setq *maxima-started* t)))
    
    (if ($file_search *maxima-initlisp*) ($load ($file_search *maxima-initlisp*)))
    (if ($file_search *maxima-initmac*) ($batchload ($file_search *maxima-initmac*)))

    (catch 'quit-to-lisp
      (in-package :maxima)
      (loop
	 do
	 (catch #+kcl si::*quit-tag*
		#+(or cmu scl sbcl openmcl lispworks) 'continue
		#-(or kcl cmu scl sbcl openmcl lispworks) nil
		(catch 'macsyma-quit
		  (continue input-stream batch-flag)
		  (format t *maxima-epilog*)
		  (bye)))))))

(defun maxima-banner ()
  (format t *maxima-prolog*)
  (format t "~&Maxima ~a http://maxima.sourceforge.net~%"
      *autoconf-version*)
  (format t (intl:gettext "using Lisp ~a ~a") (lisp-implementation-type)
      #-clisp (lisp-implementation-version)
      #+clisp (subseq (lisp-implementation-version)
	      0 (1+ (search ")" (lisp-implementation-version)))))
  (format t (intl:gettext "~%Distributed under the GNU Public License. See the file COPYING.~%"))
  (format t (intl:gettext "Dedicated to the memory of William Schelter.~%"))
  (format t (intl:gettext "The function bug_report() provides bug reporting information.~%")))

#+kcl
(si::putprop :t 'throw-macsyma-top 'si::break-command)

(defun throw-macsyma-top ()
  (throw 'macsyma-quit t))

(defmfun $writefile (x)
  (let ((msg (dribble (maxima-string x))))
    (format t "~&~A~&" msg)
    '$done))

(defvar $appendfile nil )
(defvar *appendfile-data*)

(defmfun $appendfile (name)
  (if (and (symbolp name)
	   (member (char (symbol-name name) 0) '(#\$) :test #'char=))
      (setq name (maxima-string name)))
  (if $appendfile (merror (intl:gettext "appendfile: already in appendfile, you must call closefile first.")))
  (let ((stream  (open name :direction :output
		       :if-exists :append
		       :if-does-not-exist :create)))
    (setq *appendfile-data* (list stream *terminal-io* name))

    (setq $appendfile (make-two-way-stream
		       (make-echo-stream *terminal-io* stream)
		       (make-broadcast-stream *terminal-io* stream))
	  *terminal-io* $appendfile)
    (multiple-value-bind (sec min hour day month year)
	(get-decoded-time)
      (format t (intl:gettext "~&/* Starts dribbling to ~A (~d/~d/~d, ~d:~d:~d).*/~&")
	      name year month day hour min sec))
    '$done))

(defmfun $closefile ()
  (cond ($appendfile
	 (cond ((eq $appendfile *terminal-io*)
		(format t (intl:gettext "~&/*Finished dribbling to ~A.*/~&")
			(nth 2 *appendfile-data*))
		(setq *terminal-io* (nth 1 *appendfile-data*)))
	       (t (warn "*TERMINAL-IO* was rebound while APPENDFILE is on.~%~
		   You may miss some dribble output.")))
	 (close (nth 0 *appendfile-data*))
	 (setq *appendfile-data* nil $appendfile nil))
	(t (let ((msg (dribble)))
             (format t "~&~A~&" msg))))
  '$done)

(defmfun $ed (x)
  (ed (maxima-string x)))

(defun nsubstring (x y)
  (subseq x y))

(defun filestrip (x)
  (subseq (print-invert-case (car x)) 1))

(defmspec $with_stdout (arg)
  (declare (special $file_output_append))
  (setq arg (cdr arg))
  (let ((output (meval (car arg))))
    (if (streamp output)
      (let
        ((*standard-output* output)
         (body (cdr arg))
         result)
        (dolist (v body)
          (setq result (meval* v)))
        result)
      (let*
        ((fname (namestring (maxima-string output)))
         (filespec
           (if (or (eq $file_output_append '$true)
                   (eq $file_output_append t))
             `(*standard-output* ,fname :direction :output :if-exists :append :if-does-not-exist :create)
             `(*standard-output* ,fname :direction :output :if-exists :supersede :if-does-not-exist :create))))
        (eval
          `(with-open-file ,filespec
             (let ((body ',(cdr arg)) result)
               (dolist (v body)
                 (setq result (meval* v)))
               result)))))))

(defun $sconcat (&rest x)
  (let ((ans "") )
    (dolist (v x)
      (setq ans (concatenate 'string ans
			     (cond
				   ((stringp v) v)
				   (t
				    (coerce (mstring v) 'string))))))
    ans))

(defun $system (&rest args)
  ;; If XMaxima is running, direct output from command into *SOCKET-CONNECTION*.
  ;; From what I can tell, GCL, ECL, and Clisp cannot redirect the output into an existing stream. Oh well.
  (let ((s (and (boundp '*socket-connection*) *socket-connection*))
	shell shell-opt)
    #+(or gcl ecl lispworks)
    (declare (ignore s))

    (cond ((string= *autoconf-windows* "true")
	   (setf shell "cmd") (setf shell-opt "/c"))
	  (t (setf shell "/bin/sh") (setf shell-opt "-c")))

    #+gcl (system::system (apply '$sconcat args))
    #+ecl (si:system (apply '$concat args))
    #+clisp (let ((output (ext:run-shell-command (apply '$sconcat args)
                                                 :wait t :output :stream)))
              (loop for line = (read-line output nil)
                 while line do
                   (format (or s t) "~a~%" line)))
    #+(or cmu scl) (ext:run-program shell (list shell-opt (apply '$sconcat args)) :output (or s t))
    #+allegro (excl:run-shell-command (apply '$sconcat args) :wait t :output (or s nil))
    #+sbcl (sb-ext:run-program shell
			       #+(or win32 win64) (cons shell-opt (mapcar '$sconcat args))
			       #-(or win32 win64) (list shell-opt (apply '$sconcat args))
			       :search t :output (or s t))
    #+openmcl (ccl::run-program shell
				#+windows (cons shell-opt (mapcar '$sconcat args))
				#-windows (list shell-opt (apply '$sconcat args))
				:output (or s t))
    #+abcl (extensions::run-shell-command (apply '$sconcat args) :output (or s *standard-output*))
    #+lispworks (system:run-shell-command (apply '$sconcat args) :wait t)))

(defun $room (&optional (arg nil arg-p))
  (if (and arg-p (member arg '(t nil) :test #'eq))
      (room arg)
      (room)))

(defun maxima-lisp-debugger (condition me-or-my-encapsulation)
  (declare (ignore me-or-my-encapsulation))
  (format t (intl:gettext "~&Maxima encountered a Lisp error:~%~% ~A") condition)
  (format t (intl:gettext "~&~%Automatically continuing.~%To enable the Lisp debugger set *debugger-hook* to nil.~%"))
  (throw 'return-from-debugger t))

(let ((t0-real 0) (t0-run 0)
      (float-units (float internal-time-units-per-second)))

  (defun initialize-real-and-run-time ()
    (setq t0-real (get-internal-real-time))
    (setq t0-run (get-internal-run-time)))

  (defun $absolute_real_time () (get-universal-time))

  (defun $elapsed_real_time ()
    (let ((elapsed-real-time (- (get-internal-real-time) t0-real)))
      (/ elapsed-real-time float-units)))

  (defun $elapsed_run_time ()
    (let ((elapsed-run-time (- (get-internal-run-time) t0-run)))
      (/ elapsed-run-time float-units))))

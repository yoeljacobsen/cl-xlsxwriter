(in-package #:xlsxwriter)

;;; Memorize current row and col. If only a new is provided, col is reset to 0.
;;; Convert the lisp value to the correct value
;;; String -> String
;;; Number -> DOUBLE-FLOAT
;;; NIL -> ""

(let ((current-row 0)
      (current-col 0))
  (defun worksheet-write (worksheet value &key row col format new-row)
    "Add a value to the worksheet. If row/col are not provided, add to the left"
    (let* ((actual-row (if row row
                           (if new-row (1+ current-row)current-row)))
           (actual-col (if col col
                           (if (and (null new-row)
                                    (= actual-row current-row))
                               current-col 0)))
           (actual-format (if format format (null-pointer))))
      (cond ((stringp value)
             (worksheet-write-string worksheet actual-row actual-col value actual-format))
            ((numberp value)
             (worksheet-write-number worksheet actual-row actual-col (coerce value 'double-float) actual-format))
            ((null value)
             (worksheet-write-number worksheet actual-row actual-col "" actual-format))
            (t "Unsupported value type"))
      (setf current-col (1+ actual-col))
      (setf current-row actual-row))))


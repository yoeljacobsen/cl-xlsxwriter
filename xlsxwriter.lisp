(in-package #:xlsxwriter)

;;; Make sue libxlsxwriter is installed

(define-foreign-library libxlsxwriter
  (:unix "libxlsxwriter.so")
  (t (:default "libxlsxwriter")))

(use-foreign-library libxlsxwriter)

;;; Data types

(defctype lxw-row :int32)
(defctype lxw-col :int32)
(defctype lxw-workbook :pointer)
(defctype lxw-worksheet :pointer)
(defctype lxw-format :pointer)

(defcenum lxw-error
  :LXW-NO-ERROR
  :LXW-ERROR-MEMORY-MALLOC-FAILED
  :LXW-ERROR-CREATING-XLSX-FILE
  :LXW-ERROR-CREATING-TMPFILE
  :LXW-ERROR-READING-TMPFILE
  :LXW-ERROR-ZIP-FILE-OPERATION
  :LXW-ERROR-ZIP-PARAMETER-ERROR
  :LXW-ERROR-ZIP-BAD-ZIP-FILE
  :LXW-ERROR-ZIP-INTERNAL-ERROR
  :LXW-ERROR-ZIP-FILE-ADD
  :LXW-ERROR-ZIP-CLOSE
  :LXW-ERROR-NULL-PARAMETER-IGNORED
  :LXW-ERROR-PARAMETER-VALIDATION
  :LXW-ERROR-SHEETNAME-LENGTH-EXCEEDED
  :LXW-ERROR-INVALID-SHEETNAME-CHARACTER
  :LXW-ERROR-SHEETNAME-START-END-APOSTROPHE
  :LXW-ERROR-SHEETNAME-ALREADY-USED
  :LXW-ERROR-SHEETNAME-RESERVED
  :LXW-ERROR-32-STRING-LENGTH-EXCEEDED
  :LXW-ERROR-128-STRING-LENGTH-EXCEEDED
  :LXW-ERROR-255-STRING-LENGTH-EXCEEDED
  :LXW-ERROR-MAX-STRING-LENGTH-EXCEEDED
  :LXW-ERROR-SHARED-STRING-INDEX-NOT-FOUND
  :LXW-ERROR-WORKSHEET-INDEX-OUT-OF-RANGE
  :LXW-ERROR-WORKSHEET-MAX-NUMBER-URLS-EXCEEDED
  :LXW-ERROR-IMAGE-DIMENSIONS)

;;; Functions - General

(defcfun "workbook_new" lxw-workbook (filename :string))

(defcfun "workbook_add_worksheet" lxw-worksheet
  (workbook lxw-workbook)
  (sheet-name :string))

(defcfun "worksheet_write_string" lxw-error
  (worksheet lxw-worksheet)
  (row lxw-row)
  (col lxw-col)
  (string :string)
  (format lxw-format))

(defcfun "worksheet_write_number" lxw-error
  (worksheet lxw-worksheet)
  (row lxw-row)
  (col lxw-col)
  (double :double)
  (format lxw-format))

(defcfun "workbook_close" lxw-error
  (workbook lxw-workbook))

;;; functions -formatting

(defcfun "workbook_add_format" lxw-format
  (workbook lxw-workbook))

(defcfun "format_set_bold" :void
  (format lxw-format))

;;; Helper functions

;;; Memorize current row and col and convert the lisp value to the correct  value
;;; String -> String
;;; Number -> DOUBLE-FLOAT
;;; NIL -> ""

(let ((current-row 0)
      (current-col 0))
  (defun worksheet-write (worksheet value &key row col format)
    "Add a value to the worksheet. If row/col are not provided, add to the left"
    (let ((actual-row (if row row current-row))
          (actual-col (if col col current-col))
          (actual-format (if format format (null-pointer))))
      (cond ((stringp value)
             (worksheet-write-string worksheet actual-row actual-col value actual-format))
            ((numberp value)
             (worksheet-write-number worksheet actual-row actual-col (coerce value 'double-float) actual-format))
            ((null value)
             (worksheet-write-number worksheet actual-row actual-col "" actual-format))
        (t "Unsupported value type"))
      (setf current-row actual-row)
      (setf current-col (1+ actual-col)))))

;;; Examples

(defun lxw-example (filename)
  (let* ((workbook (workbook-new filename))
         (worksheet (workbook-add-worksheet workbook "sheet1")))
    (worksheet-write-string worksheet 10 10 "Hello world" (null-pointer))
    (worksheet-write-number worksheet 11 10 42d0 (null-pointer))
    (workbook-close workbook)))

(defun test-helper (filename)
  (let* ((workbook (workbook-new filename))
         (worksheet (workbook-add-worksheet workbook "sheet1"))
         (f1 (workbook-add-format workbook)))
    (format-set-bold f1)
    (loop for val in '(1 "two" 3.0 "four" 5/6 "six")
          do (worksheet-write worksheet val :format f1))
    (workbook-close workbook)))


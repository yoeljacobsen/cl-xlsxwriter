(in-package #:xlsxwriter)

;;; Make sue libxlsxwriter is installed

(define-foreign-library libxlsxwriter
  (:unix "libxlsxwriter.so")
  (t (:default "libxlsxwriter")))

(use-foreign-library libxlsxwriter)

;;; Data types

(defctype lxw-row :int32)
(defctype lxw-col :int32)

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

;;; Functions

(defcfun "workbook_new" :pointer (filename :string))

(defcfun "workbook_add_worksheet" :pointer
  (workbook :pointer)
  (sheet-name :string))

(defcfun "worksheet_write_string" lxw-error
  (worksheet :pointer)
  (row lxw-row)
  (col lxw-col)
  (string :string)
  (format :pointer))

(defcfun "worksheet_write_number" lxw-error
  (worksheet :pointer)
  (row lxw-row)
  (col lxw-col)
  (double :double)
  (format :pointer))

(defcfun "workbook_close" lxw-error
  (workbook :pointer))

;;; Examples

(defun lxw-example (filename)
  (let* ((workbook (workbook-new filename))
         (worksheet (workbook-add-worksheet workbook "sheet1")))
    (worksheet-write-string worksheet 10 10 "Hello world" (null-pointer))
    (worksheet-write-number worksheet 11 10 42d0 (null-pointer))
    (workbook-close workbook)))

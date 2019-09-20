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
(defctype lxw-color :int32)

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

(defcenum lxw-format-alignments
  :LXW-ALIGN-NONE
  :LXW-ALIGN-LEFT
  :LXW-ALIGN-CENTER
  :LXW-ALIGN-RIGHT
  :LXW-ALIGN-FILL
  :LXW-ALIGN-JUSTIFY
  :LXW-ALIGN-CENTER-ACROSS
  :LXW-ALIGN-DISTRIBUTED
  :LXW-ALIGN-VERTICAL-TOP
  :LXW-ALIGN-VERTICAL-BOTTOM
  :LXW-ALIGN-VERTICAL-CENTER
  :LXW-ALIGN-VERTICAL-JUSTIFY
  :LXW-ALIGN-VERTICAL_DISTRIBUTED)

(defcenum lxw-format-patterns
  :LXW-PATTERN-NONE
  :LXW-PATTERN-SOLID
  :LXW-PATTERN-MEDIUM_GRAY
  :LXW-PATTERN-DARK_GRAY
  :LXW-PATTERN-LIGHT_GRAY
  :LXW-PATTERN-DARK_HORIZONTAL
  :LXW-PATTERN-DARK_VERTICAL
  :LXW-PATTERN-DARK_DOWN
  :LXW-PATTERN-DARK_UP
  :LXW-PATTERN-DARK_GRID
  :LXW-PATTERN-DARK_TRELLIS
  :LXW-PATTERN-LIGHT_HORIZONTAL
  :LXW-PATTERN-LIGHT_VERTICAL
  :LXW-PATTERN-LIGHT_DOWN
  :LXW-PATTERN-LIGHT_UPc
  :LXW-PATTERN-LIGHT_GRID
  :LXW-PATTERN-LIGHT_TRELLIS
  :LXW-PATTERN-GRAY_125
  :LXW-PATTERN-GRAY_0625)

(defcenum lxw-format-borders
  :LXW-BORDER-NONE
  :LXW-BORDER-THIN
  :LXW-BORDER-MEDIUM
  :LXW-BORDER-DASHED
  :LXW-BORDER-DOTTED
  :LXW-BORDER-THICK
  :LXW-BORDER-DOUBLE
  :LXW-BORDER-HAIR
  :LXW-BORDER-MEDIUM_DASHED
  :LXW-BORDER-DASH_DOT
  :LXW-BORDER-MEDIUM_DASH_DOT
  :LXW-BORDER-DASH_DOT_DOT
  :LXW-BORDER-MEDIUM_DASH_DOT_DOT
  :LXW-BORDER-SLANT_DASH_DOT)


;;; Constants

(defconstant LXW-COLOR-BLACK 	  #x1000000)
(defconstant LXW-COLOR-BLUE 	  #x0000FF)
(defconstant LXW-COLOR-BROWN 	  #x800000)
(defconstant LXW-COLOR-CYAN 	  #x00FFFF)
(defconstant LXW-COLOR-GRAY 	  #x808080)
(defconstant LXW-COLOR-GREEN 	  #x008000)
(defconstant LXW-COLOR-LIME 	  #x00FF00)
(defconstant LXW-COLOR-MAGENTA 	#xFF00FF)
(defconstant LXW-COLOR-NAVY 	  #x000080)
(defconstant LXW-COLOR-ORANGE 	#xFF6600)
(defconstant LXW-COLOR-PINK 	  #xFF00FF)
(defconstant LXW-COLOR-PURPLE 	#x800080)
(defconstant LXW-COLOR-RED 	    #xFF0000)
(defconstant LXW-COLOR-SILVER 	#xC0C0C0)
(defconstant LXW-COLOR-WHITE 	  #xFFFFFF)
(defconstant LXW-COLOR-YELLOW 	#xFFFF00)

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

(defcfun "format_set_font_color" :void
  (format lxw-format)
  (color lxw-color))

(defcfun "format_set_align" :void
  (format lxw-format)
  (alignment lxw-format-alignments))

(defcfun "format_set_bg_color" :void
  (format lxw-format)
  (color lxw-color))

(defcfun "format_set_fg_color" :void
  (format lxw-format)
  (color lxw-color))

(defcfun "format_set_pattern" :void
  (format lxw-format)
  (pattern lxw-format-patterns))

(defcfun "format_set_border" :void
  (format lxw-format)
  (border lxw-format-borders))

(defcfun "format_set_bottom" :void
  (format lxw-format)
  (border lxw-format-borders))

(defcfun "format_set_top" :void
  (format lxw-format)
  (border lxw-format-borders))

(defcfun "format_set_left" :void
  (format lxw-format)
  (border lxw-format-borders))

(defcfun "format_set_right" :void
  (format lxw-format)
  (border lxw-format-borders))

;;; Helper functions

;;; Memorize current row and col. If only a new is provided, col is reset to 0.
;;; Convert the lisp value to the correct value
;;; String -> String
;;; Number -> DOUBLE-FLOAT
;;; NIL -> ""

(let ((current-row 0)
      (current-col 0))
  (defun worksheet-write (worksheet value &key row col format)
    "Add a value to the worksheet. If row/col are not provided, add to the left"
    (let* ((actual-row (if row row current-row))
           (actual-col (if col col
                          (if (= actual-row current-row) current-col 0)))
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
         (f1 (workbook-add-format workbook))
         (f2 (workbook-add-format workbook)))
    (format-set-bold f1)
    (format-set-font-color f2 lxw-color-white)
    (format-set-bg-color f2 lxw-color-blue)
    (format-set-pattern f2 :lxw-pattern-solid)
    (format-set-align f2 :lxw-align-center)
    (format-set-border f2 :lxw-border-thick)
    (loop for val in '(1 "two" 3.0 "four" 5/6 "six")
          do (worksheet-write worksheet val :format f1))
    (loop for val in '(1 "two" 3.0 "four" 5/6 "six")
          do (worksheet-write worksheet val :row 3 :format f2))
    (workbook-close workbook)))


(in-package #:xlsxwriter)

;;; Make sue libxlsxwriter is installed

(define-foreign-library libxlsxwriter
  (:darwin "libxlsxwriter.dylib")
  (:unix "libxlsxwriter.so")
  (t (:default "libxlsxwriter")))

(use-foreign-library libxlsxwriter)


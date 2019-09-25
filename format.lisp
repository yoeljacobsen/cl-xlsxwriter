(in-package #:xlsxwriter)

(defcfun "format_set_font_size" :void
  (format lxw-formaT)
  (size :double))

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

(defcfun "format_set_text_wrap" :void
  (format lxw-format))


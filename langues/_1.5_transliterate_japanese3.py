# -*- coding: utf-8 -*-

import pykakasi

kakasi = pykakasi.kakasi()
kakasi.setMode("H","a") # default: Hiragana no conversion
kakasi.setMode("K","a") # default: Katakana no conversion
kakasi.setMode("J","a") # default: Japanese no conversion
kakasi.setMode("r","Hepburn") # default: use Hepburn Roman table
kakasi.setMode("C", True) # add space default: no Separator
kakasi.setMode("c", False) # capitalize default: no Capitalize
converter  = kakasi.getConverter()
result=converter.do(u'彼は負傷した')
print unicode(result)

kakasi = pykakasi.kakasi()
kakasi.setMode("H","a")
kakasi.setMode("K","a")
kakasi.setMode("J","a")
kakasi.setMode("r","Kunrei")
kakasi.setMode("E","a")
kakasi.setMode("a", None)
kakasi.setMode("C", True)
converter  = kakasi.getConverter()
result=converter.do('彼は負傷した')
print result



kakasi = pykakasi.kakasi()
kakasi.setMode("H","a")
kakasi.setMode("K","a")
kakasi.setMode("J","a")
kakasi.setMode("r","Passport")
kakasi.setMode("E","a")
kakasi.setMode("a",None)
converter  = kakasi.getConverter()
result=converter.do('彼は負傷した')
print result


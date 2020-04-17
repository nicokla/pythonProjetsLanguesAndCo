from lang_trans.arabic.buckwalter import Buckwalter

trans = Buckwalter()

print trans.transliterate(u'هذا البرنامج يعطينا نطق الحروف')
print trans.untransliterate('h*A AlbrnAmj yETynA nTq AlHrwf')

transliterateArabic('أخذت حقيبتي وغادرت.')

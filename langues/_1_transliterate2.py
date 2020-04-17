# -*- coding: UTF-8 -*-

ru2la = [
[u'а',u'a'],
[u'з',u'z'],
[u'э',u'e'],
[u'р',u'r'],
[u'т',u't'],
[u'ы',u'ы'],
[u'у',u'u'],
[u'и',u'i'],
[u'о',u'o'],
[u'п',u'p'],
[u'ш',u'sh'],
[u'с',u's'],
[u'д',u'd'],
[u'Ф',u'f'],
[u'г',u'g'],
[u'х',u'kh'],
[u'ж',u'j'],
[u'к',u'k'],
[u'л',u'l'],
[u'м',u'm'],
[u'щ',u'S'],
[u'ь',u'\''],
[u'ъ',u'\"'],
[u'в',u'v'],
[u'б',u'b'],
[u'н',u'n'],
[u'я',u'ia'],
[u'е',u'ie'],
[u'ю',u'iu'],
[u'й',u'y'],
[u'ё',u'io'],
[u'ч',u'tsh'],
[u'ц',u'ts'],
[u'А',u'A'],
[u'З',u'Z'],
[u'Э',u'E'],
[u'Р',u'R'],
[u'Т',u'T'],
[u'Ы',u'Ы'],
[u'У',u'U'],
[u'И',u'I'],
[u'О',u'O'],
[u'П',u'P'],
[u'Ш',u'Sh'],
[u'С',u'S'],
[u'Д',u'D'],
[u'ф',u'F'],
[u'Г',u'G'],
[u'Х',u'Kh'],
[u'Ж',u'J'],
[u'К',u'K'],
[u'Л',u'L'],
[u'М',u'M'],
[u'Щ',u'S'],
[u'Ь',u'\''],
[u'Ъ',u'\"'],
[u'В',u'V'],
[u'Б',u'B'],
[u'Н',u'N'],
[u'Я',u'Ia'],
[u'Е',u'Ie'],
[u'Ю',u'iu'],
[u'Й',u'y'],
[u'Ё',u'Io'],
[u'Ч',u'Tsh'],
[u'Ц',u'ts']]

he2la = [
[u'שׁ',u'sh'],
[u'שׂ',u's'],
[u'וֹ',u'o'],
[u'וּ',u'u'],
[u'א',u'\''],
[u'ז',u'z'],
[u'ע',u'3'],
[u'ר',u'r'],
[u'ת',u't'],
[u'י',u'y'],
[u'וו',u'v'],
[u'ו',u'v'],
[u'ק',u'q'],
[u'פ',u'p'],
[u'ש',u'S'],
[u'ד',u'd'],
[u'ג',u'g'],
[u'ה',u'h'],
[u'ח',u'kh'],
[u'כ',u'k'],
[u'ל',u'l'],
[u'מ',u'm'],
[u'נ',u'n'],
[u'ט',u't'],
[u'צ',u'ts'],
[u'ס',u's'],
[u'ב',u'b'],
[u'נ',u'n'],
[u'ף',u'f'],
[u'ך',u'kh'],
[u'ן',u'n'],
[u'ץ',u'ts'],
[u'ם',u'm'],
[u'ַ',u'a'],
[u'ָ',u'a'],
[u'ֵ',u'e'],
[u'ֶ',u'e'],
[u'ּ',u'u'],
[u'ִ',u'i'],
[u'ֹ',u'o'],
[u'ְ',u'']
]


ar2la=[
[u'ا',u'a'],
[u'ز',u'z'],
[u'ع',u'3'],
[u'ر',u'r'],
[u'ت',u't'],
[u'ط',u'T'],
[u'و',u'w'],
[u'ي',u'ii'],
[u'لأ',u'Al'],
[u'ق',u'q'],
[u'س',u's'],
[u'د',u'd'],
[u'ف',u'f'],
[u'ح',u'H'],
[u'ه',u'h'],
[u'ج',u'j'],
[u'ك',u'k'],
[u'ل',u'l'],
[u'م',u'm'],
[u'ص',u'S'],
[u'أ',u'a'],
[u'ب',u'b'],
[u'ن',u'n'],
[u'ى',u'aa'],
[u'غ',u'gh'],
[u'ذ',u'z'],
[u'ث',u't'],
[u'ظ',u'Z'],
[u'ُ',u'u'],
[u'ِ',u'i'],
[u'َ',u'a'],
[u'لإ',u'Al'],
[u'ش',u'sh'],
[u'خ',u'kh'],
[u'ة',u'a'],
[u'ض',u'D'],
[u'إ',u'I']
]






def transString(langue2lat, string):
    for k in langue2lat:
		#string = string.replace(k[0], unicode(k[1]))
          string=string.replace(k[0], k[1])
    return string

def transliterateRussian(s):
    return transString(ru2la, s)
    
def transliterateArabic(s):
    return transString(ar2la, s)

def transliterateHebrew(s):
    return transString(he2la, s)


transliterateArabic(u'هذا البرنامج يعطينا نطق الحروف')
transliterateRussian(u'я говорю')
print transliterateHebrew(u'שׁמע')



from transliterate import translit, get_available_language_codes

print(translit(u"Λορεμ ιψθμ δολορ σιτ αμετ", 'el', reversed=True))
print(translit(u"Λορεμ ιψθμ δολορ σιτ αμετ", reversed=True))


print(translit(u"Лорем ипсум долор сит амет", 'ru', reversed=True))
print(translit(u"Лорем ипсум долор сит амет", reversed=True))

print(translit(u'я говорю', 'ru', reversed=True))

transliterateRussian('я говорю')
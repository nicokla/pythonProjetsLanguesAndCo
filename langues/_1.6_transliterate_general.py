http://polyglot.readthedocs.io/en/latest/Transliteration.html

Thai, Amharaic, Yiddish, Hebrew, Korean




from polyglot.transliteration import Transliterator

from polyglot.downloader import downloader
print(downloader.supported_languages_table("transliteration2"))
 1. Haitian; Haitian Creole    2. Tamil                      3. Vietnamese
 4. Telugu                     5. Croatian                   6. Hungarian
 7. Thai                       8. Kannada                    9. Tagalog
10. Armenian                  11. Hebrew (modern)           12. Turkish
13. Portuguese                14. Belarusian                15. Norwegian Nynorsk
16. Norwegian                 17. Dutch                     18. Japanese
19. Albanian                  20. Bulgarian                 21. Serbian
22. Swahili                   23. Swedish                   24. French
25. Latin                     26. Czech                     27. Yiddish
28. Hindi                     29. Danish                    30. Finnish
31. German                    32. Bosnian-Croatian-Serbian  33. Slovak
34. Persian                   35. Lithuanian                36. Slovene
37. Latvian                   38. Bosnian                   39. Gujarati
40. Italian                   41. Icelandic                 42. Spanish; Castilian
43. Ukrainian                 44. Georgian                  45. Urdu
46. Indonesian                47. Marathi (Marāṭhī)         48. Korean
49. Galician                  50. Khmer                     51. Catalan; Valencian
52. Romanian, Moldavian, ...  53. Basque                    54. Macedonian
55. Russian                   56. Azerbaijani               57. Chinese
58. Estonian                  59. Welsh                     60. Arabic
61. Bengali                   62. Amharic                   63. Irish
64. Malay                     65. Afrikaans                 66. Polish
67. Greek, Modern             68. Esperanto                 69. Maltese

polyglot download embeddings2.en transliteration2.ar

from polyglot.text import Text


blob = """We will meet at eight o'clock on Thursday morning."""
text = Text(blob)

for x in text.transliterate("ar"):
  print(x)
  
  
  
  
  
  
  
  





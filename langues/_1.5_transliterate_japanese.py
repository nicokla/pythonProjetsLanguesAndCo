
#str.translate(table)
https://docs.python.org/3/library/stdtypes.html#str.translate

katakana_chart = "ァアィイゥウェエォオカガキギクグケゲコゴサザシジスズセゼソゾタダチヂッツヅテデトドナニヌネノハバパヒビピフブプヘベペホボポマミムメモャヤュユョヨラリルレロヮワヰヱヲンヴヵヶヽヾ"
hiragana_chart = "ぁあぃいぅうぇえぉおかがきぎくぐけげこごさざしじすずせぜそぞただちぢっつづてでとどなにぬねのはばぱひびぴふぶぷへべぺほぼぽまみむめもゃやゅゆょよらりるれろゎわゐゑをんゔゕゖゝゞ" 
hir2kat = str.maketrans(hiragana_chart, katakana_chart)
kat2hir  =str.maketrans(katakana_chart, hiragana_chart)


mixed = 'きゃりーぱみゅぱみゅは日本の歌手です。'
print(mixed.translate(hir2kat))
# out: キャリーパミュパミュハ日本ノ歌手デス。

# transliterate back and forth
print(mixed.translate(hir2kat).translate(kat2hir))
# out: きゃりーぱみゅぱみゅは日本の歌手です。

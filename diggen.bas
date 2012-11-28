Attribute VB_Name = "diggen"
Dim dpath As String

Function getpull(suf)

5 durk = Int(Rnd * 20)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "yank"
If durk = 1 Then durkstr = "pull"
If durk = 2 Then durkstr = "engulf"
If durk = 3 Then durkstr = "suck"
If durk = 4 Then durkstr = "slurp"
If durk = 5 Then durkstr = "ingest"

getpull = durkstr & suf

End Function

Function getburn(suf)

5 durk = Int(Rnd * 20)
If durk >= 7 Then GoTo 5

If durk = 0 Then durkstr = "burn"
If durk = 1 Then durkstr = "digest"
'If durk = 1 Then durkstr = "eat"
'If durk = 1 And suf = "ed" Then suf = "en"
If durk = 2 Then durkstr = "dissolve"
If durk = 3 Then durkstr = "liquify"
If durk = 4 Then durkstr = "melt"
If durk = 5 Then durkstr = "digest"
If durk = 6 Then durkstr = "sear"
If durk = 7 Then durkstr = ""
If durk = 8 Then durkstr = ""
If durk = 9 Then durkstr = ""
If durk = 10 Then durkstr = ""
If durk = 11 Then durkstr = ""
If durk = 12 Then durkstr = ""
If durk = 13 Then durkstr = ""
If durk = 14 Then durkstr = ""
If durk = 15 Then durkstr = ""
If durk = 16 Then durkstr = ""
If durk = 17 Then durkstr = ""
If durk = 18 Then durkstr = ""
If durk = 19 Then durkstr = ""
If durk = 20 Then durkstr = ""

getburn = durkstr & suf

End Function

Function getacid(suf)

5 durk = Int(Rnd * 20)
If durk >= 4 Then GoTo 5

If durk = 0 Then durkstr = "acid"
If durk = 1 Then durkstr = "liquid"
If durk = 2 Then durkstr = "juice"
If durk = 3 Then durkstr = "enzymes"
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getacid = durkstr & suf


End Function

Function getslither(suf)

5 durk = Int(Rnd * 20)
If durk >= 5 Then GoTo 5

If durk = 0 Then durkstr = "slither"
If durk = 1 Then durkstr = "slide"
If durk = 2 Then durkstr = "slop"
If durk = 3 Then durkstr = "squeeze"
If durk = 4 Then durkstr = "slurp"
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getslither = durkstr & suf

End Function

Function getfood(suf)

5 durk = Int(Rnd * 20)
If durk >= 17 Then GoTo 5

If durk = 0 Then durkstr = "hamburger"
If durk = 1 Then durkstr = "potato chip"
If durk = 2 Then durkstr = "sandwiche"
If durk = 3 Then durkstr = "pizza"
If durk = 4 Then durkstr = "salad"
If durk = 5 Then durkstr = "meatloaf"
If durk = 6 Then durkstr = "cookie"
If durk = 7 Then durkstr = "cheese"
If durk = 8 Then durkstr = "carrot"
If durk = 9 Then durkstr = "vegetable"
If durk = 10 Then durkstr = "fruit"
If durk = 11 Then durkstr = "apple"
If durk = 12 Then durkstr = "beef"
If durk = 13 Then durkstr = "scrambled egg"
If durk = 14 Then durkstr = "steak dinner"
If durk = 15 Then durkstr = "candy bar"
If durk = 16 Then durkstr = "steak"
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getfood = durkstr & suf

End Function

Function getdrink(suf)

5 durk = Int(Rnd * 20)
If durk >= 13 Then GoTo 5

If durk = 0 Then durkstr = "fruit juice"
If durk = 1 Then durkstr = "WATER"
If durk = 2 Then durkstr = "apple juice"
If durk = 3 Then durkstr = "orange juice"
If durk = 4 Then durkstr = "wine"
If durk = 5 Then durkstr = "milk"
If durk = 6 Then durkstr = "coke"
If durk = 7 Then durkstr = "sprite"
If durk = 8 Then durkstr = "pepsi"
If durk = 9 Then durkstr = "mountain dew"
If durk = 10 Then durkstr = "coffee"
If durk = 11 Then durkstr = "chocolate milk"
If durk = 12 Then durkstr = "lemonade"
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getdrink = durkstr & suf

End Function

Function getchunks(suf)

5 durk = Int(Rnd * 20)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "bit"
If durk = 1 Then durkstr = "chunk"
If durk = 2 Then durkstr = "hunk"
If durk = 3 Then durkstr = "blob"
If durk = 4 Then durkstr = "lump"
If durk = 5 Then durkstr = "piece"
If durk = 6 Then durkstr = "lump"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getchunks = durkstr & suf

End Function

Function getdigest(suf)

5 durk = Int(Rnd * 20)
If durk >= 7 Then GoTo 5

If durk = 0 Then durkstr = "digest"
If durk = 1 Then durkstr = "absorb"
If durk = 2 Then durkstr = "churn"
If durk = 3 Then durkstr = "crush"
If durk = 4 Then durkstr = "squeeze"
If durk = 5 Then durkstr = "dissolve"
If durk = 6 Then durkstr = "absorb"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getdigest = durkstr & suf

End Function

Function getwet(suf)

5 durk = Int(Rnd * 20)
If durk >= 14 Then GoTo 5

If durk = 0 Then durkstr = "wet"
If durk = 1 Then durkstr = "moist"
If durk = 2 Then durkstr = "sloppy"
If durk = 3 Then durkstr = "warm"
If durk = 4 Then durkstr = "humid"
If durk = 5 Then durkstr = "hot"
If durk = 6 Then durkstr = "shimmering"
If durk = 7 Then durkstr = "oozing"
If durk = 8 Then durkstr = "glistening"
If durk = 9 Then durkstr = "slick"
If durk = 10 Then durkstr = "damp"
If durk = 11 Then durkstr = "sexy"
If durk = 12 Then durkstr = "dripping"
If durk = 13 Then durkstr = "slimey"
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getwet = durkstr & suf

End Function

Function getvomit(suf)

5 durk = Int(Rnd * 20)
If durk >= 7 Then GoTo 5

If durk = 0 Then durkstr = "digestive " & getacid("s")
If durk = 1 Then durkstr = getwet("") & " digestive " & getacid("s")
If durk = 2 Then durkstr = getwet(" ") & "puke"
If durk = 3 Then durkstr = getwet(" ") & "vomit"
If durk = 4 Then durkstr = getwet("") & " gastric " & getacid("s")
If durk = 5 Then durkstr = geticky("") & " gastric " & getacid("s")
If durk = 6 Then durkstr = geticky("") & " digestive " & getacid("s")
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getvomit = durkstr & suf

End Function

Function getthroat(suf)

5 durk = Int(Rnd * 20)
If durk >= 4 Then GoTo 5

If durk = 0 Then durkstr = "esophagus"
If durk = 1 Then durkstr = "hatch"
If durk = 2 Then durkstr = "throat"
If durk = 3 Then durkstr = "gullet"
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getthroat = durkstr & suf

End Function

Function gethuge(suf)

5 durk = Int(Rnd * 20)
If durk >= 7 Then GoTo 5

If durk = 0 Then durkstr = "gaping"
If durk = 1 Then durkstr = "enourmous"
If durk = 2 Then durkstr = "cavernous"
If durk = 3 Then durkstr = "huge"
If durk = 4 Then durkstr = "gargantuan"
If durk = 5 Then durkstr = "abyssal"
If durk = 6 Then durkstr = "unending"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

gethuge = durkstr & suf


End Function

Function getswallow(suf)

5 durk = Int(Rnd * 20)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "gulp"
If durk = 1 Then durkstr = "swallow"
If durk = 2 Then durkstr = "engulf"
If durk = 3 Then durkstr = "slurp"
If durk = 4 Then durkstr = "ingest"
If durk = 5 Then durkstr = "swallow"
'If durk = 6 Then durkstr = "eat"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getswallow = durkstr & suf

End Function

Function getsexy(suf)

5 durk = Int(Rnd * 5)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "sexy"
If durk = 1 Then durkstr = "beautiful"
If durk = 2 Then durkstr = "shapely"
If durk = 3 Then durkstr = "comely"
If durk = 4 Then durkstr = "attractive"
If durk = 5 Then durkstr = "gorgeous"

If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getsexy = durkstr & suf


End Function

Function gettasty(suf)

5 durk = Int(Rnd * 5)
If durk >= 3 Then GoTo 5

If durk = 0 Then durkstr = "tasty"
If durk = 1 Then durkstr = "yummy"
If durk = 2 Then durkstr = "delicious"
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
gettasty = durkstr & suf


End Function

Function getbelly(suf)

5 durk = Int(Rnd * 20)
If durk >= 11 Then GoTo 5

If durk = 0 Then durkstr = "belly"
If durk = 1 Then durkstr = "stomach"
If durk = 2 Then durkstr = "tummy"
If durk = 3 Then durkstr = "guts"
If durk = 4 Then durkstr = "digestive tract"
If durk = 5 Then durkstr = "digestive system"
If durk = 6 Then durkstr = "innards"
If durk = 7 Then durkstr = "belly"
If durk = 8 Then durkstr = "stomach"
If durk = 9 Then durkstr = "tummy"
If durk = 10 Then durkstr = "body"
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getbelly = durkstr & suf



End Function

Function gettotal(suf)

5 durk = Int(Rnd * 20)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "totally"
If durk = 1 Then durkstr = "completely"
If durk = 2 Then durkstr = "utterly"
If durk = 3 Then durkstr = "irrevocably"
If durk = 4 Then durkstr = "entirely"
If durk = 5 Then durkstr = "fully"
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

gettotal = durkstr & suf

End Function

Function getlots(suf)

5 durk = Int(Rnd * 20)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "incredibly"
If durk = 1 Then durkstr = "unbelievably"
If durk = 2 Then durkstr = "horribly"
If durk = 3 Then durkstr = "mind-bogglingly"
If durk = 4 Then durkstr = "amazingly"
If durk = 5 Then durkstr = "indescribably"
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getlots = durkstr & suf

End Function

Function getplop(suf)

5 durk = Int(Rnd * 20)
If durk >= 5 Then GoTo 5

If durk = 0 Then durkstr = "plop"
If durk = 1 Then durkstr = "drop"
If durk = 2 Then durkstr = "pop"
If durk = 3 Then durkstr = "dump"
If durk = 4 Then durkstr = "put"
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getplop = durkstr & suf


End Function

Function getmaw(suf)

5 durk = Int(Rnd * 5)
If durk >= 3 Then GoTo 5

If durk = 0 Then durkstr = "mouth"
If durk = 1 Then durkstr = "maw"
If durk = 2 Then durkstr = "jaws"
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
'getpull = durkstr & suf

End Function

Function getpain(suf)

5 durk = Int(Rnd * 20)
If durk >= 7 Then GoTo 5

If durk = 0 Then durkstr = "horrible"
If durk = 1 Then durkstr = "terrible"
If durk = 2 Then durkstr = "agonizing"
If durk = 3 Then durkstr = "hideous"
If durk = 4 Then durkstr = "terrible"
If durk = 5 Then durkstr = "awful"
If durk = 6 Then durkstr = "unbearable"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getpain = durkstr & suf

End Function

Function getpoop(suf)

5 durk = Int(Rnd * 20)
If durk >= 9 Then GoTo 5

If durk = 0 Then durkstr = "shit"
If durk = 1 Then durkstr = "poop"
If durk = 2 Then durkstr = "crap"
If durk = 3 Then durkstr = "turd"
If durk = 4 Then durkstr = "feces"
If durk = 5 Then durkstr = "shit"
If durk = 6 Then durkstr = "excrement"
If durk = 7 Then durkstr = "poo"
If durk = 8 Then durkstr = "poop"
If durk = 9 Then durkstr = "diarhhea"
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getpoop = durkstr & suf


End Function

Function geticky(suf)

5 durk = Int(Rnd * 20)
If durk >= 11 Then GoTo 5

If durk = 0 Then durkstr = "rancid"
If durk = 1 Then durkstr = "icky"
If durk = 2 Then durkstr = "horrible"
If durk = 3 Then durkstr = "disgusting"
If durk = 4 Then durkstr = "foul"
If durk = 5 Then durkstr = "grotesque"
If durk = 6 Then durkstr = "hideous"
If durk = 7 Then durkstr = "nauseating"
If durk = 8 Then durkstr = "sickening"
If durk = 9 Then durkstr = "vile"
If durk = 10 Then durkstr = "awful"
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

geticky = durkstr & suf


End Function

Function getwrithe(suf)

5 durk = Int(Rnd * 20)
If durk >= 3 Then GoTo 5

If durk = 0 Then durkstr = "writhe"
If durk = 1 Then durkstr = "squirm"
If durk = 2 Then durkstr = "scream"
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getwrithe = durkstr & suf


End Function

Function getdrool(suf)

5 durk = Int(Rnd * 20)
If durk >= 4 Then GoTo 5

If durk = 0 Then durkstr = "drool"
If durk = 1 Then durkstr = "slobber"
If durk = 2 Then durkstr = "saliva"
If durk = 3 And suf = "s" Then suf = "tes"
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getdrool = durkstr & suf

End Function

Function getpuddle(suf)

5 durk = Int(Rnd * 5)
If durk >= 3 Then GoTo 5

If durk = 0 Then durkstr = "puddle"
If durk = 1 Then durkstr = "pool"
If durk = 2 Then durkstr = "lake"
If durk = 3 Then durkstr = ""
If durk = 4 Then durkstr = ""
If durk = 5 Then durkstr = ""

'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getpuddle = durkstr & suf

End Function

Function getslow(suf)

5 durk = Int(Rnd * 5)
If durk >= 2 Then GoTo 5

If durk = 0 Then durkstr = "gradual"
If durk = 1 Then durkstr = "slow"
'If durk = 2 Then durkstr = ""
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""

If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getslow = durkstr & suf

End Function

Function getmeal(suf)

5 durk = Int(Rnd * 5)
If durk >= 4 Then GoTo 5

If durk = 0 Then durkstr = "last meal"
If durk = 1 Then durkstr = "breakfast"
If durk = 2 Then durkstr = "lunch"
If durk = 3 Then durkstr = "dinner"
If durk = 4 Then durkstr = ""
If durk = 5 Then durkstr = ""

If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getmeal = durkstr & suf

End Function

Function getswollen(suf)

5 durk = Int(Rnd * 5)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "swollen"
If durk = 1 Then durkstr = "bulging"
If durk = 2 Then durkstr = "enourmous"
If durk = 3 Then durkstr = "distended"
If durk = 4 Then durkstr = "stuffed"
If durk = 5 Then durkstr = "gargantuan"

If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getswollen = durkstr & suf

End Function

Function gethot(suf)

5 durk = Int(Rnd * 5)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "hot"
If durk = 1 Then durkstr = "hot"
If durk = 2 Then durkstr = "muggy"
If durk = 3 Then durkstr = "humid"
If durk = 4 Then durkstr = "warm"
If durk = 5 Then durkstr = "warm"

If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
gethot = durkstr & suf


End Function

Function getmouth(suf)

5 durk = Int(Rnd * 5)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "Hot " & getdrool("") & getdrip("s") & " down the sides of her mouth and you are coated in " & getwet(" ") & getdrool(".")
If durk = 1 Then durkstr = ""
If durk = 2 Then durkstr = ""
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getmouth = durkstr & suf


End Function

Function getdrip(suf)

5 durk = Int(Rnd * 5)
If durk >= 2 Then GoTo 5

If durk = 0 Then durkstr = "drip"
If durk = 1 Then durkstr = "ooze"
If durk = 2 Then durkstr = ""
If durk = 3 Then durkstr = ""
If durk = 4 Then durkstr = ""
If durk = 5 Then durkstr = ""

If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getdrip = durkstr & suf


End Function

Function getgirlname()


5 durk = Int(Rnd * 100)
'If durk >= 75 Then GoTo 5

durkstr = ""

If durk = 0 Then durkstr = "Tracey"
If durk = 1 Then durkstr = "Stacey"
If durk = 2 Then durkstr = "Lacey"
If durk = 3 Then durkstr = "Janet"
If durk = 4 Then durkstr = "Jennifer"
If durk = 5 Then durkstr = "Judy"
If durk = 6 Then durkstr = "Trudy"
If durk = 7 Then durkstr = "Annie"
If durk = 8 Then durkstr = "Annette"
If durk = 9 Then durkstr = "Candice"
If durk = 10 Then durkstr = "Betty"
If durk = 11 Then durkstr = "Brenda"
If durk = 12 Then durkstr = "Bonnie"
If durk = 13 Then durkstr = "Brigette"
If durk = 14 Then durkstr = "Donna"
If durk = 15 Then durkstr = "Dana"
If durk = 16 Then durkstr = "Jody"
If durk = 17 Then durkstr = "Heidi"
If durk = 18 Then durkstr = "Denise"
If durk = 19 Then durkstr = "Katie"
If durk = 20 Then durkstr = "Kate"
If durk = 21 Then durkstr = "Lana"
If durk = 22 Then durkstr = "Irene"
If durk = 23 Then durkstr = "Maura"
If durk = 24 Then durkstr = "Mandy"
If durk = 25 Then durkstr = "Amanda"
If durk = 26 Then durkstr = "Allison"
If durk = 27 Then durkstr = "Kristie"
If durk = 28 Then durkstr = "Christine"
If durk = 29 Then durkstr = "Judith"
If durk = 30 Then durkstr = "Candy"
If durk = 31 Then durkstr = "Deborah"
If durk = 32 Then durkstr = "Bobbi"
If durk = 33 Then durkstr = "Evette"
If durk = 34 Then durkstr = "Eve"
If durk = 35 Then durkstr = "June"
If durk = 36 Then durkstr = "Summer"
If durk = 37 Then durkstr = "Sandy"
If durk = 38 Then durkstr = "Patricia"
If durk = 39 Then durkstr = "Miranda"
If durk = 40 Then durkstr = "Cassie"
If durk = 41 Then durkstr = "Zoe"
If durk = 42 Then durkstr = "Ruby"
If durk = 43 Then durkstr = "Trudy"
If durk = 44 Then durkstr = "Cassandra"
If durk = 45 Then durkstr = "Bethany"
If durk = 46 Then durkstr = "Lilith"
If durk = 47 Then durkstr = "Tina"
If durk = 48 Then durkstr = "Elaine"
If durk = 49 Then durkstr = "Lain"
If durk = 50 Then durkstr = "Brittany"
If durk = 51 Then durkstr = "Jean"
If durk = 52 Then durkstr = "Janice"
If durk = 53 Then durkstr = "Miranda"
If durk = 54 Then durkstr = "Alice"
If durk = 55 Then durkstr = "Catherine"
If durk = 56 Then durkstr = "Melissa"
If durk = 57 Then durkstr = "Misty"
If durk = 58 Then durkstr = "Lisa"
If durk = 59 Then durkstr = "Marissa"
If durk = 60 Then durkstr = "Autumn"
If durk = 61 Then durkstr = "Angelique"
If durk = 62 Then durkstr = "Angela"
If durk = 63 Then durkstr = "Angelina"
If durk = 64 Then durkstr = "Alita"
If durk = 65 Then durkstr = "Sophia"
If durk = 66 Then durkstr = "Tanya"
If durk = 67 Then durkstr = "Rose"
If durk = 68 Then durkstr = "Rosalyn"
If durk = 69 Then durkstr = "Trixie"
If durk = 70 Then durkstr = "Rosietta"
If durk = 71 Then durkstr = "Cherry"
If durk = 72 Then durkstr = "Cheryl"
If durk = 73 Then durkstr = "Carol"
If durk = 74 Then durkstr = "Carolin"
If durk = 75 Then durkstr = "Shelly"

If durkstr = "" Then GoTo 5
getgirlname = durkstr

End Function

Function getloaded(suf)

5 durk = Int(Rnd * 5)
If durk >= 6 Then GoTo 5

If durk = 0 Then durkstr = "swollen"
If durk = 1 Then durkstr = "bulging"
If durk = 2 Then durkstr = "loaded"
If durk = 3 Then durkstr = "loaded"
If durk = 4 Then durkstr = "stuffed"
If durk = 5 Then durkstr = "filled"

If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getloaded = durkstr & suf

End Function

Function getshitty(suf)

5 durk = Int(Rnd * 20)
If durk >= 18 Then GoTo 5

If durk = 0 Then durkstr = "wet"
If durk = 1 Then durkstr = "moist"
If durk = 2 Then durkstr = "lumpy"
If durk = 3 Then durkstr = "warm"
If durk = 4 Then durkstr = "chunky"
If durk = 5 Then durkstr = "hot"
If durk = 6 Then durkstr = "hard"
If durk = 7 Then durkstr = "firm"
If durk = 8 Then durkstr = "gooey"
If durk = 9 Then durkstr = "slick"
If durk = 10 Then durkstr = "damp"
If durk = 11 Then durkstr = "dripping"
If durk = 12 Then durkstr = "soft"
If durk = 13 Then durkstr = "slimey"
If durk = 14 Then durkstr = "squishy"
If durk = 15 Then durkstr = "greenish"
If durk = 16 Then durkstr = "dark"
If durk = 17 Then durkstr = "stinking"
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getshitty = durkstr & suf

End Function

Function getpoop2(suf)

5 durk = Int(Rnd * 20)
If durk >= 9 Then GoTo 5

If durk = 0 Then durkstr = "take a shit"
If durk = 1 Then durkstr = "take a dump"
If durk = 2 Then durkstr = "poop"
If durk = 3 Then durkstr = "relieve yourself"
If durk = 4 Then durkstr = "shit"
If durk = 5 Then durkstr = "poo"
If durk = 6 Then durkstr = "empty your bowels"
If durk = 7 Then durkstr = "unload"
If durk = 8 Then durkstr = "take a dump"
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

getpoop2 = durkstr & suf


End Function

Function getname2()

5 namec = rollgen(49)

If namec = 1 Then namee = "Dur"
If namec = 2 Then namee = "Mil"
If namec = 3 Then namee = "Lis"
If namec = 4 Then namee = "Myl"
If namec = 5 Then namee = "Tura"
If namec = 6 Then namee = "Za"
If namec = 7 Then namee = "Zana"
If namec = 8 Then namee = "Zal"
If namec = 9 Then namee = "Tri"
If namec = 10 Then namee = "Mel"
If namec = 11 Then namee = "Li"
If namec = 12 Then namee = "Lil"
If namec = 13 Then namee = "Lily"
If namec = 14 Then namee = "Ru"
If namec = 15 Then namee = "Ry"
If namec = 16 Then namee = "Ria"
If namec = 17 Then namee = "Sy"
If namec = 18 Then namee = "Cy"
If namec = 19 Then namee = "Bel"
If namec = 20 Then namee = "Mur"
If namec = 21 Then namee = "Anth"
If namec = 22 Then namee = "Ran"
If namec = 23 Then namee = "Jil"
If namec = 24 Then namee = "Ja"
If namec = 25 Then namee = "Lu"
If namec = 26 Then namee = "La"
If namec = 27 Then namee = "Na"
If namec = 28 Then namee = "Ki"
If namec = 29 Then namee = "Ka"
If namec = 30 Then namee = "Ke"
If namec = 31 Then namee = "Kat"
If namec = 32 Then namee = "Cyr"
If namec = 33 Then namee = "Cyl"
If namec = 34 Then namee = "Cyn"
If namec = 35 Then namee = "Lir"
If namec = 36 Then namee = "Lin"
If namec = 37 Then namee = "Lira"
If namec = 38 Then namee = "Elly"
If namec = 39 Then namee = "El"
If namec = 40 Then namee = "An"
If namec = 41 Then namee = "Ann"
If namec = 42 Then namee = "In"
If namec = 43 Then namee = "Il"
If namec = 44 Then namee = "Illi"
If namec = 45 Then namee = "Elli"
If namec = 46 Then namee = "Es"
If namec = 47 Then namee = "Essi"
If namec = 48 Then namee = "Al"
If namec = 0 Then namee = "Sil"

named = Int(Rnd * 45)

If named = 1 Then namef = "a"
If named = 2 Then namef = "ana"
If named = 3 Then namef = "anna"
If named = 4 Then namef = "ella"
If named = 5 Then namef = "ina"
If named = 6 Then namef = "ena"
If named = 7 Then namef = "vena"
If named = 8 Then namef = "vana"
If named = 9 Then namef = "lana"
If named = 10 Then namef = "andra"
If named = 11 Then namef = "andy"
If named = 12 Then namef = "anda"
If named = 13 Then namef = "sa"
If named = 14 Then namef = "ysa"
If named = 15 Then namef = "ette"
If named = 16 Then namef = "gette"
If named = 17 Then namef = "vela"
If named = 18 Then namef = "vala"
If named = 19 Then namef = "anca"
If named = 20 Then namef = "bel"
If named = 21 Then namef = "antha"
If named = 22 Then namef = "elle"
If named = 23 Then namef = "in"
If named = 24 Then namef = "elle"
If named = 25 Then namef = "rina"
If named = 26 Then namef = "rayna"
If named = 27 Then namef = "lina"
If named = 28 Then namef = "ie"
If named = 29 Then namef = "ranna"
If named = 30 Then namef = "isha"
If named = 31 Then namef = "ita"
If named = 32 Then namef = "lila"
If named = 33 Then namef = "lilia"
If named = 34 Then namef = "li"
If named = 35 Then namef = "ora"
If named = 36 Then namef = "lora"
If named = 37 Then namef = "lyn"
If named = 38 Then namef = "lynn"
If named = 39 Then namef = "asha"
If named = 40 Then namef = "nasha"
If named = 41 Then namef = "ly"
If named = 42 Then namef = "etta"
If named = 43 Then namef = "emma"
If named = 44 Then namef = "ella"
If named = 45 Then namef = "ara"

getname2 = namee & namef

End Function

Function gettitle()

named = Int(Rnd * 19)
If named = 0 Then nameg = "burning"
If named = 1 Then nameg = "moist"
If named = 2 Then nameg = "churning"
If named = 3 Then nameg = "thorough"
If named = 4 Then nameg = "painful"
If named = 5 Then nameg = "slithering"
If named = 6 Then nameg = "slimy"
If named = 7 Then nameg = "crushing"
If named = 8 Then nameg = "incredible"
If named = 9 Then nameg = "vile"
If named = 10 Then nameg = "stuffed"
If named = 11 Then nameg = "crowded"
If named = 12 Then nameg = "firey"
If named = 13 Then nameg = "lethal"
If named = 14 Then nameg = "overzealous"
If named = 15 Then nameg = "unyielding"
If named = 16 Then nameg = "quick-working"
If named = 17 Then nameg = "powerful"
If named = 18 Then nameg = "foul-smelling"
If named = 19 Then nameg = "rancid"
If named = 0 Then nameg = "ever-hungry"

8 named = Int(Rnd * 10)
If named < 3 Then nameh = "stomach"
If named = 3 Then nameh = "intestine"
If named = 4 Then nameh = "digestive tract"
If named = 5 Then nameh = "intestinal tract"
If named = 6 Then nameh = "guts"
If named = 7 Then nameh = "digestive organs"
If named = 8 Then nameh = "gastrointestinal tract"
If named = 9 Then nameh = "bowels"
If named = 10 Then GoTo 8

girlnm = (" of the " & nameg & " " & nameh)
'Text1.Text = ("They call her " & girlnm & ".")

gettitle = girlnm

End Function

Function getname()
zippy = rollgen(4)
getname = getgirlname
If zippy > 2 Then getname = getname2
If zippy = 2 Then getname = getname3
End Function

Function getfigure()
5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "athletic"
If durk = 1 Then durkstr = "thin"
If durk = 2 Then durkstr = "voluptuous"
If durk = 3 Then durkstr = "tight"
If durk = 4 Then durkstr = "curvaceous"
If durk = 5 Then durkstr = "shapely"
If durk = 6 Then durkstr = "compelling"
If durk = 7 Then durkstr = "perfect"
If durk = 8 Then durkstr = "luscious"
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getfigure = durkstr '& suf

End Function

Function getboobsize()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "large"
If durk = 1 Then durkstr = "large"
If durk = 2 Then durkstr = "firm"
If durk = 3 Then durkstr = "fat"
If durk = 4 Then durkstr = "small"
If durk = 5 Then durkstr = "enormous"
If durk = 6 Then durkstr = "titanic"
If durk = 7 Then durkstr = "healthy"
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getboobsize = durkstr & suf


End Function

Function getboobs()
getboobs = getboobsize & " " & getboob2 & " breasts"
End Function

Function getbreasts()
5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "breasts"
If durk = 1 Then durkstr = "breasts"
If durk = 2 Then durkstr = "boobs"
If durk = 3 Then durkstr = "hooters"
If durk = 4 Then durkstr = "chest"
If durk = 5 Then durkstr = "breasts"
If durk = 6 Then durkstr = "knockers"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getbreasts = durkstr & suf


End Function

Function getboob2()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "round"
If durk = 1 Then durkstr = "shapely"
If durk = 2 Then durkstr = "jutting"
If durk = 3 Then durkstr = "bouncing"
If durk = 4 Then durkstr = "supple"
If durk = 5 Then durkstr = "firm"
If durk = 6 Then durkstr = "pert"
If durk = 7 Then durkstr = "soft"
If durk = 8 Then durkstr = "jiggling"
If durk = 9 Then durkstr = "curving"
If durk = 10 Then durkstr = "compelling"
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getboob2 = durkstr & suf


End Function

Function gethaircol()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "red"
If durk = 1 Then durkstr = "black"
If durk = 2 Then durkstr = "jet black"
If durk = 3 Then durkstr = "brown"
If durk = 4 Then durkstr = "brunette"
If durk = 5 Then durkstr = "dark brown"
If durk = 6 Then durkstr = "sandy brown"
If durk = 7 Then durkstr = "light brown"
If durk = 8 Then durkstr = "sandy blonde"
If durk = 9 Then durkstr = "blonde"
If durk = 10 Then durkstr = "golden"
If durk = 11 Then durkstr = "auburn"
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
gethaircol = durkstr & suf


End Function

Function gethair()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "long"
If durk = 1 Then durkstr = "short"
If durk = 2 Then durkstr = "short and wavy"
If durk = 3 Then durkstr = "long and wavy"
If durk = 4 Then durkstr = "short and curly"
If durk = 5 Then durkstr = "long and curly"
If durk = 6 Then durkstr = "long"
If durk = 7 Then durkstr = "shoulder-length"
If durk = 8 Then durkstr = "pony-tailed"
If durk = 9 Then durkstr = "braided"
If durk = 10 Then durkstr = "pig-tailed"
If durk = 11 Then durkstr = "double-braided"
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
gethair = durkstr & suf


End Function

Function getmat()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "silk"
If durk = 1 Then durkstr = "cotton"
If durk = 2 Then durkstr = "polyester"
If durk = 3 Then durkstr = "leather"
If durk = 4 Then durkstr = "lace"
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getmat = durkstr & suf

End Function

Function getcolor()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "red"
If durk = 1 Then durkstr = "black"
If durk = 2 Then durkstr = "yellow"
If durk = 3 Then durkstr = "purple"
If durk = 4 Then durkstr = "black"
If durk = 5 Then durkstr = "white"
If durk = 6 Then durkstr = "grey"
If durk = 7 Then durkstr = "peach"
If durk = 8 Then durkstr = "pink"
If durk = 9 Then durkstr = "white"
If durk = 10 Then durkstr = "leopard-skin"
If durk = 11 Then durkstr = "jade"
If durk = 12 Then durkstr = "green"
If durk = 13 Then durkstr = "blue"
If durk = 14 Then durkstr = "navy"
If durk = 15 Then durkstr = "burgundy"
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getcolor = durkstr & suf

End Function

Function getclothes()

5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "bra and panties"
If durk = 1 Then durkstr = "sports bra and shorts"
If durk = 2 Then durkstr = "dress"
If durk = 3 Then durkstr = "chun-li outfit"
If durk = 4 Then durkstr = "evening dress"
If durk = 5 Then durkstr = "bikini"
If durk = 6 Then durkstr = "swimsuit"
If durk = 7 Then durkstr = "office jacket and miniskirt"
If durk = 8 Then durkstr = "sweater and skirt"
If durk = 9 Then durkstr = "miniskirt and suit"
If durk = 10 Then durkstr = "shirt and jeans"
If durk = 11 Then durkstr = "jeans and white shirt"
If durk = 12 Then durkstr = "levi jacket and pants"
If durk = 13 Then durkstr = "evening dress"
If durk = 14 Then durkstr = "catsuit"
If durk = 15 Then durkstr = "bra and panties"
If durk = 16 Then durkstr = "bra and panties"
If durk = 17 Then durkstr = "teddy"
If durk = 18 Then durkstr = "teddy"
If durk = 19 Then durkstr = "T-shirt and panties"
If durk = 20 Then durkstr = "robe"

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getclothes = durkstr & suf


End Function

Function getupto()


5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "ankles"
If durk = 1 Then durkstr = "knees"
If durk = 2 Then durkstr = "waist"
If durk = 3 Then durkstr = "stomach"
If durk = 4 Then durkstr = "elbows"
If durk = 5 Then durkstr = "chest"
If durk = 6 Then durkstr = "neck"
If durk = 7 Then durkstr = "nose"
If durk = 8 Then durkstr = "chin"
If durk = 9 Then durkstr = "shoulders"
If durk = 10 Then durkstr = "thighs"
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getupto = durkstr & suf



End Function

Function getooze()


5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "ooze"
If durk = 1 Then durkstr = "liquid"
If durk = 2 Then durkstr = "slime"
If durk = 3 Then durkstr = "mucous"
If durk = 4 Then durkstr = "goo"
If durk = 5 Then durkstr = "goop"
If durk = 6 Then durkstr = "fluid"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getooze = durkstr & suf


End Function

Function getwalk()


5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "walk"
If durk = 1 Then durkstr = "saunter"
If durk = 2 Then durkstr = "amble"
If durk = 3 Then durkstr = "strut"
If durk = 4 Then durkstr = "stride"
If durk = 5 Then durkstr = "dance"
If durk = 6 Then durkstr = "walk"
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getwalk = durkstr & suf


End Function

Function getunloads()


5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "shits"
If durk = 1 Then durkstr = "unloads"
If durk = 2 Then durkstr = "emits"
If durk = 3 Then durkstr = "drops"
If durk = 4 Then durkstr = "shits"
If durk = 5 Then durkstr = "shits"
If durk = 6 Then durkstr = "poops"
If durk = 7 Then durkstr = "dumps"
If durk = 8 Then durkstr = "excretes"
If durk = 9 Then durkstr = "drops"
If durk = 10 Then durkstr = "fires"
If durk = 11 Then durkstr = "poops"
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getunloads = durkstr & suf


End Function

Function gethappily()


5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "gradually"
If durk = 1 Then durkstr = "quickly"
If durk = 2 Then durkstr = "enthusiastically"
If durk = 3 Then durkstr = "happily"
If durk = 4 Then durkstr = "obligingly"
If durk = 5 Then durkstr = "cordially"
If durk = 6 Then durkstr = "rapidly"
If durk = 7 Then durkstr = "promptly"
If durk = 8 Then durkstr = "promptly"
If durk = 9 Then durkstr = "immediately"
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
gethappily = durkstr & suf


End Function

Function getshits()

getshits = getunloads & " a " & getbig & " load of " & getshitty(" ") & getpoop("")

End Function

Function getbig()


5 durk = Int(Rnd * 20)
If durk >= 20 Then GoTo 5

durkstr = ""
If durk = 0 Then durkstr = "big"
If durk = 1 Then durkstr = "huge"
If durk = 2 Then durkstr = "titanic"
If durk = 3 Then durkstr = "immense"
If durk = 4 Then durkstr = "gargantuan"
If durk = 5 Then durkstr = "giant"
If durk = 6 Then durkstr = "gigantic"
If durk = 7 Then durkstr = "large"
If durk = 8 Then durkstr = "enormous"
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
getbig = durkstr & suf


End Function

Function gettime(suf)


5 durk = Int(Rnd * 5)
If durk >= 4 Then GoTo 5

If durk = 0 Then durkstr = "day"
If durk = 1 Then durkstr = "hour"
If durk = 2 Then durkstr = "hour"
If durk = 3 Then durkstr = "minute"
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
gettime = durkstr & suf

End Function

Function getswear(Optional suf As String = "") As String

getswear = diggetgener("fuck", "damn", "shit", "dammit", "oh my god", "crap", "oh hell")

End Function

Function getbarf(Optional suf As String = "") As String
getbarf = diggetgener("barf", "hurl", "puke", "vomit", "throw-up")
End Function

'5 durk = Int(Rnd * 20)
'If durk >= 20 Then GoTo 5

'durkstr=""
'If durk = 0 Then durkstr = ""
'If durk = 1 Then durkstr = ""
'If durk = 2 Then durkstr = ""
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""
'If durk = 6 Then durkstr = ""
'If durk = 7 Then durkstr = ""
'If durk = 8 Then durkstr = ""
'If durk = 9 Then durkstr = ""
'If durk = 10 Then durkstr = ""
'If durk = 11 Then durkstr = ""
'If durk = 12 Then durkstr = ""
'If durk = 13 Then durkstr = ""
'If durk = 14 Then durkstr = ""
'If durk = 15 Then durkstr = ""
'If durk = 16 Then durkstr = ""
'If durk = 17 Then durkstr = ""
'If durk = 18 Then durkstr = ""
'If durk = 19 Then durkstr = ""
'If durk = 20 Then durkstr = ""

'If durkstr = "" Then GoTo 5

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
'getpull = durkstr & suf



'5 durk = Int(Rnd * 5)
'If durk >= 6 Then GoTo 5

'If durk = 0 Then durkstr = ""
'If durk = 1 Then durkstr = ""
'If durk = 2 Then durkstr = ""
'If durk = 3 Then durkstr = ""
'If durk = 4 Then durkstr = ""
'If durk = 5 Then durkstr = ""

'If suf = "es" And Right(durk, 1) = "e" Then suf = "s"
'If suf = "ed" And Right(durk, 1) = "e" Then suf = "d"
'getpull = durkstr & suf

Function getabsorb() As String

getabsorb = getgener("digest", "absorb", "dissolv")

End Function

Function getname3()

nam = getgener("R", "L", "K", "D", "T", "Tr", "S", "St", "M", "Z", "L", "N", "V", "Y", "I", "X")

For syl = 1 To rollgen(2)
    nam = nam & getgener("a", "e", "i", "o", "u", "y", "ae", "ai", "ea", "ia", "io", "ey")
    nam = nam & getgener("l", "ll", "t", "m", "r", "w", "n", "v", "c", "l", "v", "z", "k", "j", "h", "s", "ss")
Next syl

'nam = nam & getgener("ette", "a", "e", "elle", "ia", "i", "anne", "an", "a", "e", "el", "", "", "")
nam = nam & "a"

getname3 = nam

End Function

Function rollgen(ByVal damage)
damage = Int(damage)
rollgen = Int((damage - 1 + 1) * Rnd + 1)

End Function

Function geteat()
geteat = getgener("eat", "devour", "swallow", "ingest")
End Function

Function geteaten()
geteaten = getgener("eaten", "devoured", "swallowed", "ingested", "swallowed")
End Function

Function geteating()
geteating = getgener("eating", "devouring", "swallowing", "ingesting")
End Function

Function getblab()
getblab = getgener("Uh huh...", "Yeah...", "Sure...", "Whatever...", "You don't say...", "Really?", "Wow.", "Isn't that something.", "Interesting.", "Fascinating.", "Mm.", "Is that so...")
End Function

Function createwant()
cwant = getgener("$HELLO Did you want something?", _
"$HELLO What can I do for you?", _
"$HELLO Is there something you wanted?", _
"$HELLO Was there something you were interested in?", _
"$HELLO Is there some way I can help you?", _
"$HELLO Was there anything you wanted?" _
)
swaptxt cwant, "$HELLO", getgener("Anyway, ", "So, ", "So anyway, ", "Well, ")

createwant = cwant

End Function

Function getplanetname()

sletter = getgener("X", "Y", "Z", "G", "R", "Ch", "T", "Tr", "L", "R", "P", "M", "N", "Qu", "W", "V", "D", "K", "S", "L", "J")
pref = getgener("el", "ar", "ec", "il", "ez", "ak", "ar", "il", "on", "om", "el", "ec", "ar", "or", "an", "or", "om", "ax", "ed", "ay", "an")
suf1 = getgener("us", "is", "as", "isis", "ra", "rasa", "risa", "kanna", "kinna", "zac", "tira", "mena", "son", "rel", "rec", "recca", "aphel", "tiral", "kazad", "zitra", "calohn")

getplanetname = sletter & pref & suf1

End Function

Function getacidstr(amt)

If amt <= 1 Then gstr = "watery"
If amt = 2 Then gstr = "weak"
If amt = 3 Then gstr = "mild"
If amt = 4 Then gstr = "tingling"
If amt = 5 Then gstr = "burning"
If amt = 6 Then gstr = "stinging"
If amt = 7 Then gstr = "firey"
If amt = 8 Then gstr = "searing"
If amt >= 9 Then gstr = "hellish"

getacidstr = gstr

End Function

Function getchurnstr(amt)

If amt <= 1 Then gstr = "soft"
If amt = 2 Then gstr = "gentle"
If amt = 3 Then gstr = "slow"
If amt = 4 Then gstr = "steady"
If amt = 5 Then gstr = "rapid"
If amt = 6 Then gstr = "squeezing"
If amt = 7 Then gstr = "painful"
If amt = 8 Then gstr = "brutal"
If amt >= 9 Then gstr = "hellish"

getchurnstr = gstr

End Function

Function getbyamt(amt, txt1 As String, Optional txt2 As String, Optional txt3 As String, Optional txt4 As String, Optional txt5 _
                  As String, Optional txt6 As String, Optional txt7 As String, Optional txt8 As String, Optional txt9 _
                  As String, Optional txt10 As String, Optional txt11 As String, Optional txt12 As String, Optional txt13 _
                  As String, Optional txt14 As String, Optional txt15 As String, Optional txt16 As String, Optional txt17 _
                  As String, Optional txt18 As String, Optional txt19 As String, Optional txt20 As String, Optional txt21 As String) As String

amt = Int(amt)
5 If amt < 0 Then Exit Function
arollgen = amt

Select Case arollgen
    Case 1: gstr = txt1
    Case 2: gstr = txt2
    Case 3: gstr = txt3
    Case 4: gstr = txt4
    Case 5: gstr = txt5
    Case 6: gstr = txt6
    Case 7: gstr = txt7
    Case 8: gstr = txt8
    Case 9: gstr = txt9
    Case 10: gstr = txt10
    Case 11: gstr = txt11
    Case 12: gstr = txt12
    Case 13: gstr = txt13
    Case 14: gstr = txt14
    Case 15: gstr = txt15
    Case 16: gstr = txt16
    Case 17: gstr = txt17
    Case 18: gstr = txt18
    Case 19: gstr = txt19
    Case 20: gstr = txt20
    Case 21: gstr = txt21
End Select

If gstr = "" Then amt = amt - 1: GoTo 5
getbyamt = gstr
End Function

'Function getplanetname2()

'fword = getgener("Way", "Far", "Star", "Black", "Red", "White", "Grey", "Fire")
'sword=getgener("gate", "land", "cloud", "world",

'End Function

Function diggetgener(txt1 As String, Optional txt2 As String, Optional txt3 As String, Optional txt4 As String, Optional txt5 _
                  As String, Optional txt6 As String, Optional txt7 As String, Optional txt8 As String, Optional txt9 _
                  As String, Optional txt10 As String, Optional txt11 As String, Optional txt12 As String, Optional txt13 _
                  As String, Optional txt14 As String, Optional txt15 As String, Optional txt16 As String, Optional txt17 _
                  As String, Optional txt18 As String, Optional txt19 As String, Optional txt20 As String, Optional txt21 As String) As String

5 arollgen = roll(21)

Select Case arollgen
    Case 1: gstr = txt1
    Case 2: gstr = txt2
    Case 3: gstr = txt3
    Case 4: gstr = txt4
    Case 5: gstr = txt5
    Case 6: gstr = txt6
    Case 7: gstr = txt7
    Case 8: gstr = txt8
    Case 9: gstr = txt9
    Case 10: gstr = txt10
    Case 11: gstr = txt11
    Case 12: gstr = txt12
    Case 13: gstr = txt13
    Case 14: gstr = txt14
    Case 15: gstr = txt15
    Case 16: gstr = txt16
    Case 17: gstr = txt17
    Case 18: gstr = txt18
    Case 19: gstr = txt19
    Case 20: gstr = txt20
    Case 21: gstr = txt21
End Select

If gstr = "" Then GoTo 5
diggetgener = gstr
End Function

Function gettaunt() As String

aroll = roll(34)

Select Case aroll
    Case 1: dstr = "Enjoying the scenery, dear?"
    Case 2: dstr = "urp."
    Case 3: dstr = "belch."
    Case 4: dstr = "Yum. nothing like a meal that squirms after you eat it."
    Case 5: dstr = "Look on the bright side. I bet you're nutritious!"
    Case 6: dstr = "You're gonna be fat on my " & getbreasts & " now."
    Case 7: dstr = "Now you know what " & getfood("s") & " feel like, eh?"
    Case 8: dstr = "Don't worry, I've set aside a special place in my intestines for you."
    Case 9: dstr = "Ohhhhh that feels so gooooood!!"
    Case 10: dstr = "Get used to it down there, " & diggetgener("DINNER.", "LUNCH.", "FOOD.", "BREAKFAST.")
    Case 11: dstr = "Enjoy your defeat."
    Case 12: dstr = "Oh, listen to my stomach groan! It obviously likes you!"
    Case 13: dstr = getwrithe("") & " all you want.  You're in my " & getbelly("") & " for good now."
    Case 14: dstr = "Deal with it, honey. You'll be " & getpoop("") & " in a few hours."
    Case 15: dstr = "MMmmm, you were " & gettasty(".")
    Case 16: dstr = "Soon you'll be digested and you'll become part of my " & getsexy("") & " body!"
    Case 17: dstr = "You act like you've never been swallowed whole before."
    Case 18: dstr = geticky(", isn't it?")
    Case 19: dstr = "Enjoy my " & getbelly("!")
    Case 20: dstr = "Do you like it in my " & getbelly("?")
    Case 21: dstr = "You're in my " & getbelly("") & " and you're never gonna get out!"
    Case 22: dstr = "Oh, stop squirming.  It can't be that bad."
    Case 23: dstr = "Are you enjoying being digested?"
    Case 24: dstr = "Gateway to the digestive tract, baby.  Enjoy it."
    Case 25: dstr = "Just think.  In a few hours you'll be part of my " & getboobs & "."
    Case 26: dstr = "Give it up.  You're gonna be fat on my " & getgener("thighs", "hips", "ass", getboobs) & " in a few hours."
    Case 27: dstr = "How do you like being " & getgener("a meal", "digested", "digested alive", "digested like a little " & getfood(""), "my " & getmeal("")) & ", you little " & getbadname & "."
    Case 28: dstr = "In a little while you're going to be " & getpoop("") & ", you little " & getbadname & ".  " & getgener("How do you like that?", "How does that make you feel, you little " & getbadname("?"), "And I'm gonna " & getpoop("") & " you out into my toilet like the little " & getpoop("") & " you are.", "Isn't that exciting?", "Doesn't that thought feel good?")
    Case 29: dstr = "Does it hurt yet?"
    Case 30: dstr = "I can't believe you haven't passed out yet.  You're a tough little " & getsolidfood(".")
    Case 31: dstr = "I hope you like " & getgener("stomach ", "gastric ", "digestive ") & getgener("juices", "acids", "liquids") & ", you little " & getbadname(".")
    Case 32: dstr = "I am going to digest every cell in your body!"
    Case 33: dstr = "I hope you enjoy getting to know my " & getgener("breakfast.", "lunch.", "dinner.")
    Case 34: dstr = "At least you're good for one thing.  *burp!*"
End Select

gettaunt = dstr
End Function

Function getsolidfood(Optional suf = "")

getsolidfood = getgener("potato chip", "sandwich", "hamburger", "meatball", "snack", "celery stick", "slab of beef", "bran muffin", "chocolate chip cookie", "banana") & suf

End Function

Function getbadname(Optional suf = "")

getbadname = getgener("asshole", "prick", "fucker", "insignificant fucker", "shit", "ass licker", "cock sucker", "bastard", "cunt", "shit eater", "loser", "piece of shit") & suf

End Function

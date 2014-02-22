Attribute VB_Name = "Rand"
Option Explicit
Const MODULE_NAME$ = "Randomizer"

Public Const PEARLS_MAX& = 55

Private Enum RandomizerErrors    ' you may make this Public for tests
    ErrorBase = vbObjectError + 513    ' you may adjust this minimum
    NotInitted
    AlreadyInitted
    ' add error numbers here
End Enum

Private Type ErrorHolder            '
    HasError As Boolean             '  temp storage for errors
    Source As String                '
    Number As RandomizerErrors    '
    Description As String
End Type
Private mError As ErrorHolder

Private mRandomized As Boolean _
      , mAlphabetsFilled As Boolean

Private mVowels$(), mConsonants$()






'   I sent a letter to the fish,                I said it very loud and clear,
'   I told them, "This is what I wish."         I went and shouted in his ear.
'
'   The little fishes of the sea,               But he was very stiff and proud,
'   They sent an answer back to me.             He said "You needn't shout so loud."
'
'   The little fishes' answer was               And he was very proud and stiff,
'   "We cannot do it, sir, because..."          He said "I'll go and wake them if..."
'
'   I sent a letter back to say                 I took a kettle from the shelf,
'   It would be better to obey.                 I went to wake them up myself.
'
'   But someone came to me and said             But when I found the door was locked
'   "The little fishes are in bed."             I pulled and pushed and kicked and knocked,
'
'   I said to him, and I said it plain          And when I found the door was shut,
'   "Then you must wake them up again."         I tried to turn the handle, But...
'
'
'           "Is that all?" asked Alice.
'           "That is all." said Humpty Dumpty. "Goodbye."




'   Matter will be damaged in direct proportion to its value.

'   Logic is a systematic method of coming to the wrong conclusion with confidence.


'   Florence Flask was ... dressing for the opera when she turned to her
'   husband and screamed, "Erlenmeyer!  My joules!  Someone has stolen my
'   joules!"
'
'   "Now, now, my dear," replied her husband, "keep your balance and reflux
'   a moment.  Perhaps they're mislead."
'
'   "No, I know they're stolen," cried Florence.  "I remember putting them
'   in my burette ... We must call a copper."
'
'   Erlenmeyer did so, and the flatfoot who turned up, one Sherlock Ohms,
'   said the outrage looked like the work of an arch-criminal by the name
'   of Lawrence Ium.
'
'   "We must be careful -- he's a free radical, ultraviolet, and
'   dangerous.  His girlfriend is a chlorine at the Palladium.  Maybe I can
'   catch him there."  With that, he jumped on his carbon cycle in an
'   activated state and sped off along the reaction pathway ...
'           -- Daniel B. Murphy, "Precipitations"
'

'   If at first you don't succeed, redefine success.

'   I've been on a diet for two weeks and all I've lost is two weeks.
'       -- Totie Fields


'   With listening comes wisdom, with speaking repentance.


'   The error of youth is to believe that intelligence is a substitute for
'   experience, while the error of age is to believe experience is a substitute
'   for intelligence.
'           -- Lyman Bryson


'   What we Are is God's gift to us.
'   What we Become is our gift to God.


'   Matter cannot be created or destroyed, nor can it be returned without a receipt.



'   The Worst Bank Robbery
'
'       In August 1975 three men were on their way in to rob the Royal Bank of
'   Scotland at Rothesay, when they got stuck in the revolving doors.  They
'   had to be helped free by the staff and, after thanking everyone,
'   sheepishly left the building.
'
'       A few minutes later they returned and announced their intention of
'   robbing the bank, but none of the staff believed them.  When they demanded
'   5,000 pounds in cash, the head cashier laughed at them, convinced that it
'   was a practical joke.
'
'       Then one of the men jumped over the counter, but fell to the floor
'   clutching his ankle.  The other two tried to make their getaway, but got
'   trapped in the revolving doors again.




'   Shell to DOS... Come in DOS, do you copy?  Shell to DOS...




'   user, n.:
'       The word computer professionals use when they mean "idiot."
'           -- Dave Barry, "Claw Your Way to the Top"
'
'   [I always thought "computer professional" was the phrase hackers used
'    when they meant "idiot."  Ed.]




'   Laws of Computer Programming:
'       (1) Any given program, when running, is obsolete.
'       (2) Any given program costs more and takes longer.
'       (3) If a program is useful, it will have to be changed.
'       (4) If a program is useless, it will have to be documented.
'       (5) Any given program will expand to fill all available memory.
'       (6) The value of a program is proportional the weight of its output.
'       (7) Program complexity grows until it exceeds the capability of
'               the programmer who must maintain it.


'   If builders built buildings the way programmers wrote programs,
'   then the first woodpecker to come along would destroy civilization.

'

Public Property Get mNumber(Optional lowrBound As Long = 1 _
                          , Optional upprBound As Long = 32767 _
                          ) As Long
    On Error GoTo ErrH             ' Default max is largest Integer.
                                   '
    If Not mRandomized Then                   '
        Call Randomize
        mRandomized = True
    End If
    
    mNumber = Int((upprBound - lowrBound + 1) * Rnd + lowrBound)
    
ErrH: Blame "mNumber"
End Property

Public Function mNumbers(Optional minNumbrs As Long = 1 _
                       , Optional ByVal maxNumbrs As Long = 4 _
                       , Optional lowrBound As Long = 1 _
                       , Optional upprBound As Long = 32767 _
                       ) As Long()
    Dim i&, numbrs&, nn&()
    On Error GoTo Cleanup
    
    If maxNumbrs < minNumbrs Then maxNumbrs = minNumbrs
    
    numbrs = Rand.mNumber(minNumbrs, maxNumbrs)
    
    ReDim nn(numbrs - 1)
    For i = 0 To UBound(nn)
        nn(i) = Rand.mNumber(lowrBound, upprBound)
    Next i
    
    mNumbers = nn
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase nn
  LoadError "mNumbers" ', "details of error"
End Function


'Public Function mLetter() As String
'    On Error GoTo ErrH
'
'    If Not mAlphabetsFilled Then
'        Call FillAlphabets
'        mAlphabetsFilled = True
'    End If
'
'    If Rand.mBoolean Then
'        mLetter = mConsonants(Rand.mNumber(0, UBound(mConsonants)))
'    Else
'        mLetter = mVowels(Rand.mNumber(0, UBound(mVowels)))
'    End If
'
'ErrH: Blame "mLetter"
'End Function

Public Function mFullName() As String
    Dim firstNme$, midl$, surnme$
    On Error GoTo ErrH
    
    If Not mAlphabetsFilled Then
        Call FillAlphabets
        mAlphabetsFilled = True
    End If
    
    firstNme = Rand.mPhrase(1, 3)
    midl = UCase$(mConsonants(Rand.mNumber(0, UBound(mConsonants)))) & "."
    surnme = Rand.mPhrase(1, 2)
    
    mFullName = Rand.mTitle _
        & " " & firstNme _
        & " " & midl _
        & " " & surnme
    
ErrH: Blame "mFullName"
End Function


Public Property Get mSyllable() As String   '  Concatenate a random vowel
    On Error GoTo ErrH                      '   and a random consonant
                                            '    to get a random syllable.
    If Not mAlphabetsFilled Then
        Call FillAlphabets
        mAlphabetsFilled = True
    End If
    
    mSyllable = mConsonants(Rand.mNumber(0, UBound(mConsonants))) _
                  & mVowels(Rand.mNumber(0, UBound(mVowels)))
    
ErrH: Blame "mSyllable" ', "details of error"
End Property



Public Property Get mSyllables(Optional minSylables As Long = 1 _
                             , Optional ByVal maxSylables As Long = 4 _
                             , Optional capitalize1stLetter As Boolean = True _
                             ) As String()
    Dim i&, sylables&, ss$()
    On Error GoTo Cleanup
    
    If maxSylables < minSylables Then maxSylables = minSylables
    
    sylables = Rand.mNumber(minSylables, maxSylables)
    
    ReDim ss(sylables - 1)
    For i = 0 To UBound(ss)
        ss(i) = Rand.mSyllable()
    Next i
    
    mSyllables = ss
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "mSyllables" ', "details of error"
End Property


Public Property Get mWord(Optional minSylables As Long = 1 _
                        , Optional ByVal maxSylables As Long = 4 _
                        , Optional minCharactrs As Long = 4 _
                        , Optional capitalize1stLetter As Boolean = True _
                        ) As String
    On Error GoTo ErrH
    
    
    mWord = Join(Rand.mSyllables(minSylables, maxSylables _
                               , capitalize1stLetter) _
               , vbNullString)
    
    ' if less than minimum chars, generate a new one
    If Len(mWord) < minCharactrs Then mWord _
        = Rand.mWord(minSylables, maxSylables, minCharactrs, False)
    
    If capitalize1stLetter Then mWord = ProperCase(mWord)
    
ErrH: Blame "mWord" ', "details of error"
End Property


Public Function mWords(Optional minmumWords As Long = 2 _
                     , Optional ByVal maxmumWords As Long = 4 _
                     , Optional capitalize1stLetters As Boolean = True _
                     , Optional minSylables As Long = 1 _
                     , Optional ByVal maxSylables As Long = 4 _
                     , Optional minCharactrs As Long = 4 _
                     ) As String()
    Dim i&, wrdCount&, ss$()
    On Error GoTo Cleanup
    
    If maxmumWords < minmumWords Then maxmumWords = minmumWords
    
    wrdCount = Rand.mNumber(minmumWords, maxmumWords)
    
    ReDim ss(wrdCount - 1)
    For i = 0 To UBound(ss)
        ss(i) = Rand.mWord(minSylables, maxSylables _
                         , minCharactrs, capitalize1stLetters)
    Next i
    
    mWords = ss
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "mWords" ', "details of error"
End Function


Public Property Get mPhrase(Optional minmumWords As Long = 2 _
                          , Optional ByVal maxmumWords As Long = 4 _
                          , Optional delimitr As String = " " _
                          , Optional capitalize1stLetters As Boolean = True _
                          , Optional minSylables As Long = 1 _
                          , Optional ByVal maxSylables As Long = 4 _
                          , Optional minCharactrs As Long = 4 _
                          ) As String
    On Error GoTo ErrH
    
    mPhrase = Join(Rand.mWords(minmumWords, maxmumWords, capitalize1stLetters _
                             , minSylables, maxSylables, minCharactrs) _
                 , delimitr)
    
ErrH: Blame "mPhrase" ', "details of error"
End Property


Public Property Get mPhrases(Optional minmumPhrases As Long = 3 _
                           , Optional ByVal maxmumPhrases As Long = 14 _
                           , Optional minmumWords As Long = 2 _
                           , Optional ByVal maxmumWords As Long = 4 _
                           , Optional delimitr As String = " " _
                           , Optional capitalize1stLetters As Boolean = True _
                           , Optional minSylables As Long = 1 _
                           , Optional ByVal maxSylables As Long = 4 _
                           , Optional minCharactrs As Long = 4 _
                           ) As String()
    Dim i&, phrseCount&, ss$()
    On Error GoTo Cleanup
    
    If maxmumPhrases < minmumPhrases Then maxmumPhrases = minmumPhrases
    
    phrseCount = Rand.mNumber(minmumPhrases, maxmumPhrases)
    
    ReDim ss(phrseCount - 1)
    For i = 0 To UBound(ss)
        ss(i) = Rand.mPhrase(minmumWords, maxmumWords _
                           , delimitr, capitalize1stLetters _
                           , minSylables, maxSylables, minCharactrs)
    Next i
    
    mPhrases = ss
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase ss
  LoadError "mPhrases" ', "details of error"
End Property


Public Property Get mTitle() As String
    Dim c(), s$()
    On Error GoTo Cleanup
    
    ReDim c(17, 1): c(0, 1) = 5:   c(0, 0) = "Ms."
                    c(1, 1) = 5:   c(1, 0) = "Mrs."
                    c(2, 1) = 5:   c(2, 0) = "Mr."
                    
                    c(3, 1) = 1:   c(3, 0) = "St."
                    c(4, 1) = 1:   c(4, 0) = "Pope"
                    c(5, 1) = 3:   c(5, 0) = "Msgr."
                    c(6, 1) = 5:   c(6, 0) = "Rev.Fr."
                    c(7, 1) = 5:   c(7, 0) = "Fr."
                    c(8, 1) = 5:   c(8, 0) = "Bro."
                    c(9, 1) = 3:   c(9, 0) = "Sister"
                    
                   c(10, 1) = 3:  c(10, 0) = "Judge"
                   c(11, 1) = 5:  c(11, 0) = "Atty."
                   c(12, 1) = 5:  c(12, 0) = "Dr."
                   c(13, 1) = 5:  c(13, 0) = "Engr."
                   
                   c(14, 1) = 1:  c(14, 0) = "Pres."
                   c(15, 1) = 2:  c(15, 0) = "Sen."
                   
                   c(16, 1) = 3:  c(16, 0) = "Gen."
                   c(17, 1) = 5:  c(17, 0) = "Capt."
    
    s = ExpandArray(c)
    
    mTitle = s(Rand.mNumber(0, UBound(s)))
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase c, s
  LoadError "mTitle" ', "details of error"
End Property


Public Property Get mItem(ParamArray poolOfItems() As Variant _
                        ) As Variant
    On Error GoTo ErrH
    
    mItem = poolOfItems(Rand.mNumber(0, UBound(poolOfItems)))
    
ErrH: Blame "mItem"
End Property


Public Property Get mBoolean() As Boolean
    On Error GoTo ErrH
    
    mBoolean = Rand.mItem(True, False)
    
ErrH: Blame "mBoolean"
End Property


Public Function mPearl(Optional ByRef srcCaption As Variant _
                     , Optional ByVal specificPointNumbr As Long = -1 _
                     ) As String
    Dim pt&(), chaptr$, src$
    On Error GoTo Cleanup
    
    If specificPointNumbr = -1 Then
        ReDim pt(1 To PEARLS_MAX):   pt(1) = 3
                                     pt(2) = 3
                                     pt(3) = 3
                                     pt(4) = 3
                                     pt(5) = 3
                                     pt(6) = 3
                                     pt(7) = 3
                                     pt(8) = 3
                                     pt(9) = 3
                                    pt(10) = 3
                                    pt(11) = 3
                                    pt(12) = 3
                                    pt(13) = 3
                                    pt(14) = 3
                                    pt(15) = 3
                                    pt(16) = 3
                                    pt(17) = 3
                                    pt(18) = 3
                                    pt(19) = 3
                                    pt(20) = 3
                                    pt(21) = 3
                                    pt(22) = 3
                                    pt(23) = 3
                                    pt(24) = 3
                                    pt(25) = 3
                                    pt(26) = 3
                                    pt(27) = 3
                                    pt(28) = 3
                                    pt(29) = 3
                                    pt(30) = 3
                                    pt(31) = 3
                                    pt(32) = 3
                                    pt(33) = 3
                                    pt(34) = 3
                                    pt(35) = 3
                                    pt(36) = 3
                                    pt(37) = 3
                                    pt(38) = 3
                                    pt(39) = 3
                                    pt(40) = 3
                                    pt(41) = 3
                                    pt(42) = 3
                                    pt(43) = 3
                                    pt(44) = 3
                                    pt(45) = 3
                                    pt(46) = 3
                                    pt(47) = 3
                                    pt(48) = 3
                                    pt(49) = 3
                                    pt(50) = 3
                                    pt(51) = 3
                                    pt(52) = 3
                                    pt(53) = 3
                                    pt(54) = 3
                                    pt(55) = 3
        
        Call ExpandArrayLng(pt)
        specificPointNumbr = pt(Rand.mNumber(1, UBound(pt)))
    End If
    
    Select Case specificPointNumbr
        
        Case 1: chaptr = "Character"
            mPearl = "Don’t let your life be sterile. Be useful. Blaze a trail. Shine forth with the light of your faith and of your love." & vbCrLf _
                   & "With your apostolic life wipe out the slimy and filthy mark left by the impure sowers of hatred. And light up all the ways of the earth with the fire of Christ that you carry in your heart."
            
        Case 2: chaptr = "Character"
            mPearl = "May your behavior and your conversation be such that everyone who sees or hears you can say: This man reads the life of Jesus Christ."
            
        Case 3: chaptr = "Character"
            mPearl = "Maturity. Stop making faces and acting up like a child! Your bearing ought to reflect the peace and order in your soul."
            
        Case 4: chaptr = "Character"
            mPearl = "Don’t say, “That’s the way I am—it’s my character.” It’s your lack of character. Esto vir!—Be a man!"
            
        Case 5: chaptr = "Character"
            mPearl = "Get used to saying No."
            
        Case 6: chaptr = "Character"
            mPearl = "Turn your back on the deceiver when he whispers in your ear, “Why complicate your life?”"
            
        Case 7: chaptr = "Character"
            mPearl = "Don’t have a “small town” outlook. Enlarge your heart until it becomes universal—“catholic.”" & vbCrLf _
                   & "Don’t fly like a barnyard hen when you can soar like an eagle."
            
        Case 8: chaptr = "Character"
            mPearl = "Serenity. Why lose your temper if by losing it you offend God, you trouble your neighbor, you give yourself a bad time… and in the end you have to set things aright, anyway?"
            
        Case 9: chaptr = "Character"
            mPearl = "What you have just said, say it in another tone, without anger, and what you say will have more force… and above all, you won’t offend God."
            
        Case 10: chaptr = "Character"
            mPearl = "Never reprimand anyone while you feel provoked over a fault that has been committed. Wait until the next day, or even longer. Then make your remonstrance calmly and with a purified intention. You’ll gain more with an affectionate word than you ever would from three hours of quarreling. Control your temper."
            
        Case 11: chaptr = "Character"
            mPearl = "Will-power. Energy. Example. What has to be done is done… without wavering… without worrying about what others think…" & vbCrLf _
                   & "Otherwise, Cisneros would not have been Cisneros; nor Teresa of Ahumada, St. Teresa; nor Iñigo of Loyola, St. Ignatius." & vbCrLf _
                   & "God and daring! “Regnare Christum volumus!”—“We want Christ to reign!”"
            
        Case 12: chaptr = "Character"
            mPearl = "Let obstacles only make you bigger. The grace of our Lord will not be lacking: “Inter medium montium pertransibunt aquae!”—“Through the very midst of the mountains the waters shall pass.” You will pass through mountains!" & vbCrLf _
                   & "What does it matter that you have to curtail your activity for the moment, if later, like a spring which has been compressed, you’ll advance much farther than you ever dreamed?"
            
        Case 13: chaptr = "Character"
            mPearl = "Get rid of those useless thoughts which are at best a waste of time."
            
        Case 14: chaptr = "Character"
            mPearl = "Don’t waste your energy and your time—which belong to God—throwing stones at the dogs that bark at you on the way. Ignore them."
            
        Case 15: chaptr = "Character"
            mPearl = "Don’t put off your work until tomorrow."
            
        Case 16: chaptr = "Character"
            mPearl = "Give in? Be just commonplace? You, a sheep-like follower? You were born to be a leader!" & vbCrLf _
                   & "Among us there is no place for the lukewarm. Humble yourself, and Christ will kindle in you again the fire of love."
            
        Case 17: chaptr = "Character"
            mPearl = "Don’t succumb to that disease of character whose symptoms are a general lack of seriousness, unsteadiness in action and speech, foolishness—in a word, frivolity." & vbCrLf _
                   & "And that frivolity, mind you, which makes your plans so void—“so filled with emptiness”—will make of you a lifeless and useless dummy, unless you react in time—not tomorrow, but now!"
            
        Case 18: chaptr = "Character"
            mPearl = "You go on being worldly, frivolous, and giddy because you are a coward. What is it, if not cowardice, to refuse to face yourself?"
            
        Case 19: chaptr = "Character"
            mPearl = "Will-power. A very important quality. Don’t disregard the little things, which are really never futile or trivial. For by the constant practice of repeated self-denial in little things, with God’s grace you will increase in strength and manliness of character. In that way you’ll first become master of yourself, and then a guide and a leader: to compel, to urge, to draw others with your example and with your word and with your knowledge and with your power."
            
        Case 20: chaptr = "Character"
            mPearl = "You clash with the character of one person or another… It has to be that way—you are not a dollar bill to be liked by everyone." & vbCrLf _
                   & "Besides, without those clashes which arise in dealing with your neighbors, how could you ever lose the sharp corners, the edges—imperfections and defects of your character—and acquire the order, the smoothness, and the firm mildness of charity, of perfection?" & vbCrLf _
                   & "If your character and that of those around you were soft and sweet like marshmallows, you would never become a saint."
            
        Case 21: chaptr = "Character"
            mPearl = "Excuses. You’ll never lack them if you want to avoid your duties. What a lot of rationalizing!" & vbCrLf _
                   & "Don’t stop to think about excuses. Get rid of them and do what you should."
            
        Case 22: chaptr = "Character"
            mPearl = "Be firm! Be strong! Be a man! And then… be an angel!"
            
        Case 23: chaptr = "Character"
            mPearl = "You say you can’t do more? Couldn’t it be… that you can’t do less?"
            
        Case 24: chaptr = "Character"
            mPearl = "You are ambitious: for knowledge… for leadership. You want to be daring. Good. Fine. But let it be for Christ, for Love."
            
        Case 25: chaptr = "Character"
            mPearl = "Don’t argue. Arguments usually bring no light because the light is smothered by emotion."
            
        Case 26: chaptr = "Character"
            mPearl = "Matrimony is a holy sacrament. When the time comes for you to receive it, ask your spiritual director or your confessor to suggest an appropriate book. Then you’ll be better prepared to bear worthily the burdens of a home."
            
        Case 27: chaptr = "Character"
            mPearl = "Do you laugh because I tell you that you have a “vocation to marriage”? Well, you have just that—a vocation." & vbCrLf _
                   & "Commend yourself to Saint Raphael that he may keep you pure, as he did Tobias, until the end of the way."
            
        Case 28: chaptr = "Character"
            mPearl = "Marriage is for the rank and file, not for the officers of Christ’s army. For, unlike food, which is necessary for every individual, procreation is necessary only for the species, and individuals can dispense with it." & vbCrLf _
                   & "A desire to have children? Behind us we shall leave children—many children… and a lasting trail of light, if we sacrifice the selfishness of the flesh."
            
        Case 29: chaptr = "Character"
            mPearl = "The limited and pitiful happiness of the selfish man, who withdraws into his shell, his ivory tower… is not difficult to attain in this world. But that happiness of the selfish is not lasting." & vbCrLf _
                   & "For this false semblance of Heaven are you going to forsake the Joy of Glory without end?"
            
        Case 30: chaptr = "Character"
            mPearl = "You’re shrewd. But don’t tell me you are young. Youth gives all it can—it gives itself without reserve."
            
        Case 31: chaptr = "Character"
            mPearl = "Selfish! You always looking out for yourself." & vbCrLf _
                   & "You seem unable to feel the brotherhood of Christ. In others you don’t see brothers; you see stepping-stones." & vbCrLf _
                   & "I can foresee your complete failure. And when you are down, you’ll expect others to treat you with the charity you’re unwilling to show them."
            
        Case 32: chaptr = "Character"
            mPearl = "You’ll never be a leader if you see others only as stepping-stones to get ahead. You’ll be a leader if you are ambitious for the salvation of all souls." & vbCrLf _
                   & "You can’t live with your back turned on everyone; you have to be eager to make others happy."
            
        Case 33: chaptr = "Character"
            mPearl = "You never want “to get to the bottom of things.” At times, because of politeness. Other times—most times—because you fear hurting yourself. Sometimes again, because you fear hurting others. But always because of fear!" & vbCrLf _
                   & "With that fear of digging for the truth you’ll never be a man of good judgment."
            
        Case 34: chaptr = "Character"
            mPearl = "Don’t be afraid of the truth, even though the truth may mean your death."
            
        Case 35: chaptr = "Character"
            mPearl = "There are many pretty terms I don’t like: you call cowardice “prudence”. Your “prudence” gives an opportunity to those enemies of God, without any ideas in their heads, to pass themselves off as scholars, and so reach positions that they never should attain."
            
        Case 36: chaptr = "Character"
            mPearl = "Yes, that abuse can be eradicated. It’s a lack of character to let it continue as something hopeless—without any possible remedy." & vbCrLf _
                   & "Don’t evade your duty. Do it in a forthright way, even though others may not."
            
        Case 37: chaptr = "Character"
            mPearl = "You have, as they say, “the gift of gab”. But in spite of all your talk, you can’t get me to justify—by calling it “providential”—what has no justification."
            
        Case 38: chaptr = "Character"
            mPearl = "Can it be true (I just can’t believe it!) that on earth there are no men—only bellies?"
            
        Case 39: chaptr = "Character"
            mPearl = "“Pray that I may never be satisfied with what is easy”, you say. I’ve already prayed. Now it is up to you to carry out that fine resolution."
            
        Case 40: chaptr = "Character"
            mPearl = "Faith, joy, optimism. But not the folly of closing your eyes to reality."
            
        Case 41: chaptr = "Character"
            mPearl = "What a sublime way of carrying on with your empty follies, and what a way of getting somewhere in the world: rising, always rising simply by “weighing little”, by having nothing inside—neither in your head nor in your heart!"
            
        Case 42: chaptr = "Character"
            mPearl = "Why those variations in your character? When are you going to apply your will to something? Drop that craze for laying cornerstones, and finish at least one of your projects."
            
        Case 43: chaptr = "Character"
            mPearl = "Don’t be so touchy. The least thing offends you. People have to weigh their words to talk to you even about the most trivial matter." & vbCrLf _
                   & "Don’t feel hurt if I tell you that you are… unbearable. Unless you change, you’ll never be of any use."
            
        Case 44: chaptr = "Character"
            mPearl = "Use the polite excuse that Christian charity and good manners require. But then, keep on going with holy shamelessness, without stopping until you have reached the summit in the fulfillment of your duty."
            
        Case 45: chaptr = "Character"
            mPearl = "Why feel hurt by the unjust things people say of you? You would be even worse, if God ever left you." & vbCrLf _
                   & "Keep on doing good, and shrug your shoulders."
            
        Case 46: chaptr = "Character"
            mPearl = "Don’t you think that equality, as many people understand it, is synonymous with injustice?"
            
        Case 47: chaptr = "Character"
            mPearl = "That pose and those important airs don’t fit you well. It’s obvious that they’re false. At least, try not to use them either with God, or with your director, or with your brothers; and then there will be between them and you one barrier less."
            
        Case 48: chaptr = "Character"
            mPearl = "You lack character. What a mania for interfering in everything! You are bent on being the salt of every dish. And you won’t mind if I speak clearly—you haven’t the qualities of salt: you can’t be dissolved and pass unnoticed, as salt does." & vbCrLf _
                   & "You lack a spirit of sacrifice. And you abound in a spirit of curiosity and ostentation."
            
        Case 49: chaptr = "Character"
            mPearl = "Keep quiet. Don’t be “babyish”, a caricature of a child, a tattle-tale, a trouble-maker, a squealer. With your stories and tales you have chilled the warm glow of charity; you couldn’t have done more harm. And if by any chance you—your wagging tongue—have shaken down the strong walls of other people’s perseverance, your own perseverance ceases to be a grace of God. It has become a treacherous instrument of the enemy."
            
        Case 50: chaptr = "Character"
            mPearl = "You’re curious and inquisitive, prying and nosey. Aren’t you ashamed that even in your defects you are not much of a man? Be a man, and instead of poking into other people’s lives, get to know what you really are yourself."
            
        Case 51: chaptr = "Character"
            mPearl = "Your manly spirit—simple and straightforward—is crushed when you find yourself entangled in gossip and scandalous talk. You don’t understand how it could happen, and you never wished to be involved in it anyway. Suffer the humiliation that such talk causes you, and let the experience urge you to greater discretion."
            
        Case 52: chaptr = "Character"
            mPearl = "When you must judge others, why put into your criticism the bitterness of your own failures?"
            
        Case 53: chaptr = "Character"
            mPearl = "That critical spirit—granted you mean well—should never be directed toward the apostolate in which you work nor toward your brothers. In your supernatural undertakings that critical spirit—forgive me for saying it—can do a lot of harm. For when you get involved in judging the work of others, you are not doing anything constructive. Really you have no right to judge, even if you have the highest possible motives, as I admit. And with your negative attitude you hold up the progress of others." & vbCrLf _
                   & "“Then,” you ask worriedly, “my critical spirit, which is the keynote of my character…?”" & vbCrLf _
                   & "Listen. I’ll set your mind at ease. Take pen and paper. Write down simply and confidently—yes, and briefly—what is worrying you. Give the note to your superior, and don’t think any more about it. He is in charge and has the grace of state. He will file the note… or will throw it in the waste-basket. And since your criticism is not gossip and you do it for the highest motives, it’s all the same to you."
            
        Case 54: chaptr = "Character"
            mPearl = "Conform? It is a word found only in the vocabulary of those (“You might as well conform,” they say) who have no will to fight—the lazy, the cunning, the cowardly—because they know they are defeated before they start."
            
        Case 55: chaptr = "Character"
            mPearl = "Man, listen! Even though you may be like a child—and you really are one in the eyes of God—be a little less naive: don’t put your brothers “on the spot” before strangers."
        
            
        Case Else: Err.Raise specificPointNumbr, , "Invalid [pt] number: [" & specificPointNumbr & "]."
    End Select
    
    src = "The Way, """ _
        & chaptr & """, #" & specificPointNumbr & ", St. Josemaría Escrivá"
    
    
    '  if no param,
    '   - append caption to quote
    '
    If IsMissing(srcCaption) Then
        mPearl = mPearl & vbCrLf & vbTab & "-- " & src
    Else
        srcCaption = src
    End If
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase pt
  LoadError "mPearl" ', "details of error"
End Function





' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Constructors
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub FillAlphabets()
    Dim c()
    On Error GoTo Cleanup
    
    ReDim c(5, 1):  c(0, 0) = "a":  c(0, 1) = 2     ' col(x,1) = probability
                    c(1, 0) = "e":  c(1, 1) = 2     '
                    c(2, 0) = "i":  c(2, 1) = 2     '  Higher value increases
                    c(3, 0) = "o":  c(3, 1) = 2     '   chance of being picked.
                    c(4, 0) = "u":  c(4, 1) = 2     '
                    c(5, 0) = "y":  c(5, 1) = 1     '
                    
    mVowels = ExpandArray(c)
    
    
    ReDim c(36, 1): c(0, 0) = "a":   c(0, 1) = 1
                    c(1, 0) = "b":   c(1, 1) = 3
                    c(2, 0) = "c":   c(2, 1) = 3
                    c(3, 0) = "cz":  c(3, 1) = 1
                    c(4, 0) = "ch":  c(4, 1) = 2
                    c(5, 0) = "d":   c(5, 1) = 3
                    c(6, 0) = "e":   c(6, 1) = 1
                    c(7, 0) = "f":   c(7, 1) = 3
                    c(8, 0) = "ff":  c(8, 1) = 1
                    c(9, 0) = "g":   c(9, 1) = 3
                   c(10, 0) = "h":  c(10, 1) = 3
                   c(11, 0) = "i":  c(11, 1) = 1
                   c(12, 0) = "j":  c(12, 1) = 2
                   c(13, 0) = "k":  c(13, 1) = 3
                   c(14, 0) = "l":  c(14, 1) = 3
                   c(15, 0) = "m":  c(15, 1) = 3
                   c(16, 0) = "mm": c(16, 1) = 1
                   c(17, 0) = "n":  c(17, 1) = 3
                   c(18, 0) = "o":  c(18, 1) = 1
                   c(19, 0) = "ny": c(19, 1) = 1
                   c(20, 0) = "ng": c(20, 1) = 1
                   c(21, 0) = "p":  c(21, 1) = 3
                   c(22, 0) = "ph": c(22, 1) = 2
                   c(23, 0) = "q":  c(23, 1) = 1
                   c(24, 0) = "qu": c(24, 1) = 2
                   c(25, 0) = "r":  c(25, 1) = 3
                   c(26, 0) = "s":  c(26, 1) = 3
                   c(27, 0) = "sh": c(27, 1) = 2
                   c(28, 0) = "t":  c(28, 1) = 3
                   c(29, 0) = "tt": c(29, 1) = 1
                   c(30, 0) = "th": c(30, 1) = 2
                   c(31, 0) = "u":  c(31, 1) = 1
                   c(32, 0) = "v":  c(32, 1) = 2
                   c(33, 0) = "w":  c(33, 1) = 2
                   c(34, 0) = "x":  c(34, 1) = 1
                   c(35, 0) = "y":  c(35, 1) = 2
                   c(36, 0) = "z":  c(36, 1) = 2
    
    mConsonants = ExpandArray(c)
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase c
  LoadError "FillAlphabets" ', "details of error"
End Sub







' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Private Utilities
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Function ExpandArray(twoD As Variant) As String()
    Dim i&, j&, k&, itmCount&, tmp1D$()
    On Error GoTo Cleanup
    
    For i = 0 To UBound(twoD, 1)
        itmCount = itmCount + twoD(i, 1)
    Next i
    
    ReDim tmp1D(itmCount - 1)
    
    For i = 0 To UBound(twoD, 1)
        For j = 1 To twoD(i, 1)
            tmp1D(k) = twoD(i, 0)
            k = k + 1
        Next j
    Next i
    
    ExpandArray = tmp1D
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase tmp1D
  LoadError "ExpandArray" ', "details of error"
End Function

Private Sub ExpandArrayLng(ByRef lng1D() As Long)
    Dim i&, j&, k&, itmCount&, orig1D&()
    On Error GoTo Cleanup
    
    For i = 1 To UBound(lng1D)
        itmCount = itmCount + lng1D(i)
    Next i
    
    orig1D = lng1D
    ReDim lng1D(1 To itmCount)
    k = 1
    
    For i = 1 To UBound(orig1D)
        For j = 1 To orig1D(i)
            lng1D(k) = i
            k = k + 1
        Next j
    Next i
    
Cleanup:    SaveError
            'Set someObj = Nothing
            Erase orig1D
  LoadError "ExpandArrayLng" ', "details of error"
End Sub


Private Function ProperCase(strText As String) As String    ' Francesco Balena
    Dim i&                                                  '  http://www.devx.com/vb2themax/Tip/18549
    
    ' prepare the result
    ProperCase = StrConv(strText, vbProperCase)
    
    ' restore all those characters that were capitalized
    For i = 1 To Len(strText)
        Select Case Asc(Mid$(strText, i, 1))
            Case 65 To 90   ' A-Z
                Mid$(ProperCase, i, 1) = Mid$(strText, i, 1)
        End Select
    Next
End Function








' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'
'    Error Handlers
'
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

Private Sub ErrorIf(errCondition As Boolean _
                  , errorMsg As String _
                  , Optional errorNumbr As RandomizerErrors = -1 _
                  )
    If errCondition Then Err.Raise errorNumbr, MODULE_NAME, errorMsg
End Sub

Private Sub Blame(ByVal currntProcedure As String _
                , Optional ByVal errorDescrption As String _
                , Optional ByVal errorNumbr As RandomizerErrors = -1 _
                )
    Call SaveError
    Call LoadError(currntProcedure, errorDescrption, errorNumbr)
End Sub

Private Sub SaveError()
    With mError
        If Err Then
            .HasError = True
            .Description = Err.Description
            .Number = Err.Number
            .Source = Err.Source
            
        Else
            .HasError = False
            .Description = vbNullString
            .Number = 0
            .Source = vbNullString
        End If
    End With
    Err.Clear
End Sub

Private Sub LoadError(ByVal currntProcedure As String _
                    , Optional ByVal errorDescrption As String _
                    , Optional ByVal errorNumbr As RandomizerErrors = -1 _
                    )
    With mError
        If Not .HasError Then Exit Sub
            
        If LenB(errorDescrption) = 0 Then
            errorDescrption = .Description
        Else
            errorDescrption = .Description & vbCrLf & errorDescrption
        End If
        
        currntProcedure = MODULE_NAME & "." & currntProcedure & "()"

        If errorNumbr = -1 Then errorNumbr = .Number
        
        Select Case errorNumbr
            Case NotInitted
                errorDescrption = MODULE_NAME & " not initted." & vbCrLf _
                               & "Please call " & MODULE_NAME _
                               & ".Init() before " & currntProcedure & "."
            
            Case Else
                errorDescrption = currntProcedure & vbCrLf & errorDescrption
        End Select

        Err.Raise errorNumbr, .Source, errorDescrption
            
    End With
End Sub

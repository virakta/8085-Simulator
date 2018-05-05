Attribute VB_Name = "Inital"
Option Explicit
Public SModeX As Long
Public SModeY As Long
Public MaxX As Long
Public MaxY As Long
Public SModePH As Long
Public SModePW As Long
Public THeight As Long
Public TWidth As Long
Public sx As Long
Public sy As Long
Public NextPage As Integer
Public Line1 As Integer
Sub InitialiseDataTable()

InSet(1).OpCode = "CE"
InSet(1).Nemo = "ACI "
InSet(1).ByteCount = 2

InSet(2).OpCode = "8F"
InSet(2).Nemo = "ADC A "
InSet(2).ByteCount = 1

InSet(3).OpCode = "88"
InSet(3).Nemo = "ADC B "
InSet(3).ByteCount = 1

InSet(4).OpCode = "89"
InSet(4).Nemo = "ADC C "
InSet(4).ByteCount = 1

InSet(5).OpCode = "8A"
InSet(5).Nemo = "ADC D "
InSet(5).ByteCount = 1

InSet(6).OpCode = "8B"
InSet(6).Nemo = "ADC E "
InSet(6).ByteCount = 1

InSet(7).OpCode = "8C"
InSet(7).Nemo = "ADC H "
InSet(7).ByteCount = 1

InSet(8).OpCode = "8D"
InSet(8).Nemo = "ADC L"
InSet(8).ByteCount = 1

InSet(9).OpCode = "8E"
InSet(9).Nemo = "ADC M"
InSet(9).ByteCount = 1

InSet(10).OpCode = "87"
InSet(10).Nemo = "ADD A"
InSet(10).ByteCount = 1

InSet(11).OpCode = "80"
InSet(11).Nemo = "ADD B"
InSet(11).ByteCount = 1

InSet(12).OpCode = "81"
InSet(12).Nemo = "ADD C"
InSet(12).ByteCount = 1

InSet(13).OpCode = "82"
InSet(13).Nemo = "ADD D"
InSet(13).ByteCount = 1

InSet(14).OpCode = "83"
InSet(14).Nemo = "ADD E"
InSet(14).ByteCount = 1

InSet(15).OpCode = "84"
InSet(15).Nemo = "ADD H"
InSet(15).ByteCount = 1

InSet(16).OpCode = "85"
InSet(16).Nemo = "ADD L"
InSet(16).ByteCount = 1

InSet(17).OpCode = "86"
InSet(17).Nemo = "ADD M"
InSet(17).ByteCount = 1

InSet(18).OpCode = "C6"
InSet(18).Nemo = "ADI "
InSet(18).ByteCount = 2

InSet(19).OpCode = "A7"
InSet(19).Nemo = "ANA A"
InSet(19).ByteCount = 1

InSet(20).OpCode = "A0"
InSet(20).Nemo = "ANA B"
InSet(20).ByteCount = 1

InSet(21).OpCode = "A1"
InSet(21).Nemo = "ANA C"
InSet(21).ByteCount = 1

InSet(22).OpCode = "A2"
InSet(22).Nemo = "ANA D"
InSet(22).ByteCount = 1

InSet(23).OpCode = "A3"
InSet(23).Nemo = "ANA E"
InSet(23).ByteCount = 1

InSet(24).OpCode = "A4"
InSet(24).Nemo = "ANA H"
InSet(24).ByteCount = 1

InSet(25).OpCode = "A5"
InSet(25).Nemo = "ANA L"
InSet(25).ByteCount = 1

InSet(26).OpCode = "A6"
InSet(26).Nemo = "ANA M"
InSet(26).ByteCount = 1

InSet(27).OpCode = "E6"
InSet(27).Nemo = "ANI "
InSet(27).ByteCount = 2

InSet(28).OpCode = "CD"
InSet(28).Nemo = "CALL "
InSet(28).ByteCount = 3


InSet(29).OpCode = "DC"
InSet(29).Nemo = "CC "
InSet(29).ByteCount = 3

InSet(30).OpCode = "FC"
InSet(30).Nemo = "CM "
InSet(30).ByteCount = 3

InSet(31).OpCode = "2F"
InSet(31).Nemo = "CMA "
InSet(31).ByteCount = 1

InSet(32).OpCode = "3F"
InSet(32).Nemo = "CMC "
InSet(32).ByteCount = 1

InSet(33).OpCode = "BF"
InSet(33).Nemo = "CMP A"
InSet(33).ByteCount = 1

InSet(34).OpCode = "B8"
InSet(34).Nemo = "CMP B"
InSet(34).ByteCount = 1

InSet(35).OpCode = "B9"
InSet(35).Nemo = "CMP C"
InSet(35).ByteCount = 1

InSet(36).OpCode = "BA"
InSet(36).Nemo = "CMP D"
InSet(36).ByteCount = 1

InSet(37).OpCode = "BB"
InSet(37).Nemo = "CMP E"
InSet(37).ByteCount = 1

InSet(38).OpCode = "BC"
InSet(38).Nemo = "CMP H"
InSet(38).ByteCount = 1

InSet(39).OpCode = "BD"
InSet(39).Nemo = "CMP L"
InSet(39).ByteCount = 1

InSet(40).OpCode = "BE"
InSet(40).Nemo = "CMP M"
InSet(40).ByteCount = 1

InSet(41).OpCode = "D4"
InSet(41).Nemo = "CNC "
InSet(41).ByteCount = 3

InSet(42).OpCode = "C4"
InSet(42).Nemo = "CNZ "
InSet(42).ByteCount = 3

InSet(43).OpCode = "F4"
InSet(43).Nemo = "CP "
InSet(43).ByteCount = 3

InSet(44).OpCode = "EC"
InSet(44).Nemo = "CPE "
InSet(44).ByteCount = 3

InSet(45).OpCode = "FE"
InSet(45).Nemo = "CPI "
InSet(45).ByteCount = 2

InSet(46).OpCode = "E4"
InSet(46).Nemo = "CPO "
InSet(46).ByteCount = 3

InSet(47).OpCode = "CC"
InSet(47).Nemo = "CZ "
InSet(47).ByteCount = 3

InSet(48).OpCode = "27"
InSet(48).Nemo = "DAA "
InSet(48).ByteCount = 1

InSet(49).OpCode = "09"
InSet(49).Nemo = "DAD B"
InSet(49).ByteCount = 1

InSet(50).OpCode = "19"
InSet(50).Nemo = "DAD D"
InSet(50).ByteCount = 1

InSet(51).OpCode = "29"
InSet(51).Nemo = "DAD H"
InSet(51).ByteCount = 1

InSet(52).OpCode = "39"
InSet(52).Nemo = "DAD SP"
InSet(52).ByteCount = 1

InSet(53).OpCode = "3D"
InSet(53).Nemo = "DCR A"
InSet(53).ByteCount = 1

InSet(54).OpCode = "05"
InSet(54).Nemo = "DCR B"
InSet(54).ByteCount = 1

InSet(55).OpCode = "0D"
InSet(55).Nemo = "DCR C"
InSet(55).ByteCount = 1

InSet(56).OpCode = "15"
InSet(56).Nemo = "DCR D"
InSet(56).ByteCount = 1

InSet(57).OpCode = "1D"
InSet(57).Nemo = "DCR E"
InSet(57).ByteCount = 1

InSet(58).OpCode = "25"
InSet(58).Nemo = "DCR H"
InSet(58).ByteCount = 1

InSet(59).OpCode = "2D"
InSet(59).Nemo = "DCR L"
InSet(59).ByteCount = 1

InSet(60).OpCode = "35"
InSet(60).Nemo = "DCR M"
InSet(60).ByteCount = 1

InSet(61).OpCode = "0B"
InSet(61).Nemo = "DCX B"
InSet(61).ByteCount = 1

InSet(62).OpCode = "1B"
InSet(62).Nemo = "DCX D"
InSet(62).ByteCount = 1

InSet(63).OpCode = "2B"
InSet(63).Nemo = "DCX H"
InSet(63).ByteCount = 1

InSet(64).OpCode = "3B"
InSet(64).Nemo = "DCX SP"
InSet(64).ByteCount = 1

InSet(65).OpCode = "F3"
InSet(65).Nemo = "DI "
InSet(65).ByteCount = 1

InSet(66).OpCode = "FB"
InSet(66).Nemo = "EI "
InSet(66).ByteCount = 1



InSet(67).OpCode = "76"
InSet(67).Nemo = "HLT "
InSet(67).ByteCount = 1

InSet(68).OpCode = "DB"
InSet(68).Nemo = "IN "
InSet(68).ByteCount = 2

InSet(69).OpCode = "3C"
InSet(69).Nemo = "INR A"
InSet(69).ByteCount = 1

InSet(70).OpCode = "04"
InSet(70).Nemo = "INR B"
InSet(70).ByteCount = 1

InSet(71).OpCode = "0C"
InSet(71).Nemo = "INR C"
InSet(71).ByteCount = 1

InSet(72).OpCode = "14"
InSet(72).Nemo = "INR D"
InSet(72).ByteCount = 1

InSet(73).OpCode = "1C"
InSet(73).Nemo = "INR E"
InSet(73).ByteCount = 1

InSet(74).OpCode = "24"
InSet(74).Nemo = "INR H"
InSet(74).ByteCount = 1

InSet(75).OpCode = "2C"
InSet(75).Nemo = "INR L"
InSet(75).ByteCount = 1

InSet(76).OpCode = "34"
InSet(76).Nemo = "INR M"
InSet(76).ByteCount = 1

InSet(77).OpCode = "03"
InSet(77).Nemo = "INX B"
InSet(77).ByteCount = 1

InSet(78).OpCode = "13"
InSet(78).Nemo = "INX D"
InSet(78).ByteCount = 1

InSet(79).OpCode = "23"
InSet(79).Nemo = "INX H"
InSet(79).ByteCount = 1



InSet(80).OpCode = "33"
InSet(80).Nemo = "INX SP"
InSet(80).ByteCount = 1

InSet(81).OpCode = "DA"
InSet(81).Nemo = "JC "
InSet(81).ByteCount = 3

InSet(82).OpCode = "FA"
InSet(82).Nemo = "JM "
InSet(82).ByteCount = 3

InSet(83).OpCode = "C3"
InSet(83).Nemo = "JMP "
InSet(83).ByteCount = 3

InSet(84).OpCode = "D2"
InSet(84).Nemo = "JNC "
InSet(84).ByteCount = 3



InSet(85).OpCode = "C2"
InSet(85).Nemo = "JNZ "
InSet(85).ByteCount = 3

InSet(86).OpCode = "F2"
InSet(86).Nemo = "JP "
InSet(86).ByteCount = 3

InSet(87).OpCode = "EA"
InSet(87).Nemo = "JPE "
InSet(87).ByteCount = 3

InSet(88).OpCode = "E2"
InSet(88).Nemo = "JPO "
InSet(88).ByteCount = 3

InSet(89).OpCode = "CA"
InSet(89).Nemo = "JZ "
InSet(89).ByteCount = 3



InSet(90).OpCode = "3A"
InSet(90).Nemo = "LDA "
InSet(90).ByteCount = 3


InSet(91).OpCode = "0A"
InSet(91).Nemo = "LDAX B "
InSet(91).ByteCount = 1

InSet(92).OpCode = "1A"
InSet(92).Nemo = "LDAX D "
InSet(92).ByteCount = 1

InSet(93).OpCode = "2A"
InSet(93).Nemo = "LHLD "
InSet(93).ByteCount = 3

InSet(94).OpCode = "01"
InSet(94).Nemo = "LXI B "
InSet(94).ByteCount = 3

InSet(95).OpCode = "11"
InSet(95).Nemo = "LXI D "
InSet(95).ByteCount = 3

InSet(96).OpCode = "21"
InSet(96).Nemo = "LXI H "
InSet(96).ByteCount = 3

InSet(97).OpCode = "31"
InSet(97).Nemo = "LXI SP "
InSet(97).ByteCount = 3

InSet(98).OpCode = "7F"
InSet(98).Nemo = "MOV A,A"
InSet(98).ByteCount = 1


InSet(99).OpCode = "78"
InSet(99).Nemo = "MOV A,B"
InSet(99).ByteCount = 1

InSet(100).OpCode = "79"
InSet(100).Nemo = "MOV A,C"
InSet(100).ByteCount = 1

InSet(101).OpCode = "7A"
InSet(101).Nemo = "MOV A,D"
InSet(101).ByteCount = 1

InSet(102).OpCode = "7B"
InSet(102).Nemo = "MOV A,E"
InSet(102).ByteCount = 1

InSet(103).OpCode = "7C"
InSet(103).Nemo = "MOV A,H"
InSet(103).ByteCount = 1

InSet(104).OpCode = "7D"
InSet(104).Nemo = "MOV A,L"
InSet(104).ByteCount = 1

InSet(105).OpCode = "7E"
InSet(105).Nemo = "MOV A,M"
InSet(105).ByteCount = 1

InSet(106).OpCode = "47"
InSet(106).Nemo = "MOV B,A"
InSet(106).ByteCount = 1

InSet(107).OpCode = "40"
InSet(107).Nemo = "MOV B,B"
InSet(107).ByteCount = 1

InSet(108).OpCode = "41"
InSet(108).Nemo = "MOV B,C"
InSet(108).ByteCount = 1

InSet(109).OpCode = "42"
InSet(109).Nemo = "MOV B,D"
InSet(109).ByteCount = 1

InSet(110).OpCode = "43"
InSet(110).Nemo = "MOV B,E"
InSet(110).ByteCount = 1

InSet(111).OpCode = "44"
InSet(111).Nemo = "MOV B,H"
InSet(111).ByteCount = 1

InSet(112).OpCode = "45"
InSet(112).Nemo = "MOV B,L"
InSet(112).ByteCount = 1

InSet(113).OpCode = "46"
InSet(113).Nemo = "MOV B,M"
InSet(113).ByteCount = 1


InSet(114).OpCode = "4F"
InSet(114).Nemo = "MOV C,A"
InSet(114).ByteCount = 1

InSet(115).OpCode = "48"
InSet(115).Nemo = "MOV C,B"
InSet(115).ByteCount = 1

InSet(116).OpCode = "49"
InSet(116).Nemo = "MOV C,C"
InSet(116).ByteCount = 1

InSet(117).OpCode = "4A"
InSet(117).Nemo = "MOV C,D"
InSet(117).ByteCount = 1

InSet(118).OpCode = "4B"
InSet(118).Nemo = "MOV C,E"
InSet(118).ByteCount = 1

InSet(119).OpCode = "4C"
InSet(119).Nemo = "MOV C,H"
InSet(119).ByteCount = 1

InSet(120).OpCode = "4D"
InSet(120).Nemo = "MOV C,L"
InSet(120).ByteCount = 1

InSet(121).OpCode = "4E"
InSet(121).Nemo = "MOV C,M"
InSet(121).ByteCount = 1

InSet(122).OpCode = "57"
InSet(122).Nemo = "MOV D,A"
InSet(122).ByteCount = 1

InSet(123).OpCode = "50"
InSet(123).Nemo = "MOV D,B"
InSet(123).ByteCount = 1

InSet(124).OpCode = "51"
InSet(124).Nemo = "MOV D,C"
InSet(124).ByteCount = 1


InSet(125).OpCode = "52"
InSet(125).Nemo = "MOV D,D"
InSet(125).ByteCount = 1


InSet(126).OpCode = "53"
InSet(126).Nemo = "MOV D,E"
InSet(126).ByteCount = 1


InSet(127).OpCode = "54"
InSet(127).Nemo = "MOV D,H"
InSet(127).ByteCount = 1

InSet(128).OpCode = "55"
InSet(128).Nemo = "MOV D,L"
InSet(128).ByteCount = 1

InSet(129).OpCode = "56"
InSet(129).Nemo = "MOV D,M"
InSet(129).ByteCount = 1

InSet(130).OpCode = "5F"
InSet(130).Nemo = "MOV E,A"
InSet(130).ByteCount = 1

InSet(131).OpCode = "58"
InSet(131).Nemo = "MOV E,B"
InSet(131).ByteCount = 1

InSet(132).OpCode = "59"
InSet(132).Nemo = "MOV C,E"
InSet(132).ByteCount = 1

InSet(133).OpCode = "5A"
InSet(133).Nemo = "MOV E,D"
InSet(133).ByteCount = 1

InSet(134).OpCode = "5B"
InSet(134).Nemo = "MOV E,E"
InSet(134).ByteCount = 1

InSet(135).OpCode = "5C"
InSet(135).Nemo = "MOV E,H"
InSet(135).ByteCount = 1

InSet(136).OpCode = "5D"
InSet(136).Nemo = "MOV E,L"
InSet(136).ByteCount = 1

InSet(137).OpCode = "5E"
InSet(137).Nemo = "MOV E,M"
InSet(137).ByteCount = 1

InSet(138).OpCode = "67"
InSet(138).Nemo = "MOV H,A"
InSet(138).ByteCount = 1

InSet(139).OpCode = "60"
InSet(139).Nemo = "MOV H,B"
InSet(139).ByteCount = 1
InSet(140).OpCode = "61"
InSet(140).Nemo = "MOV H,C"
InSet(140).ByteCount = 1

InSet(141).OpCode = "62"
InSet(141).Nemo = "MOV H,D"
InSet(141).ByteCount = 1

InSet(142).OpCode = "63"
InSet(142).Nemo = "MOV H,E"
InSet(142).ByteCount = 1

InSet(143).OpCode = "64"
InSet(143).Nemo = "MOV H,H"
InSet(143).ByteCount = 1

InSet(144).OpCode = "65"
InSet(144).Nemo = "MOV H,L"
InSet(144).ByteCount = 1

InSet(145).OpCode = "66"
InSet(145).Nemo = "MOV H,M"
InSet(145).ByteCount = 1

InSet(146).OpCode = "6F"
InSet(146).Nemo = "MOV L,A"
InSet(146).ByteCount = 1

InSet(147).OpCode = "68"
InSet(147).Nemo = "MOV L,B"
InSet(147).ByteCount = 1

InSet(148).OpCode = "69"
InSet(148).Nemo = "MOV L,C"
InSet(148).ByteCount = 1

InSet(149).OpCode = "6A"
InSet(149).Nemo = "MOV L,D"
InSet(149).ByteCount = 1


InSet(150).OpCode = "6B"
InSet(150).Nemo = "MOV L,E"
InSet(150).ByteCount = 1

InSet(151).OpCode = "6C"
InSet(151).Nemo = "MOV L,H"
InSet(151).ByteCount = 1

InSet(152).OpCode = "6D"
InSet(152).Nemo = "MOV L,L"
InSet(152).ByteCount = 1

InSet(153).OpCode = "6E"
InSet(153).Nemo = "MOV L,M"
InSet(153).ByteCount = 1

InSet(154).OpCode = "77"
InSet(154).Nemo = "MOV M,A"
InSet(154).ByteCount = 1

InSet(155).OpCode = "70"
InSet(155).Nemo = "MOV M,B"
InSet(155).ByteCount = 1

InSet(156).OpCode = "71"
InSet(156).Nemo = "MOV M,C"
InSet(156).ByteCount = 1

InSet(157).OpCode = "72"
InSet(157).Nemo = "MOV M,D"
InSet(157).ByteCount = 1

InSet(158).OpCode = "73"
InSet(158).Nemo = "MOV M,E"
InSet(158).ByteCount = 1
InSet(159).OpCode = "74"
InSet(159).Nemo = "MOV M,H"
InSet(159).ByteCount = 1
InSet(160).OpCode = "75"
InSet(160).Nemo = "MOV M,L"
InSet(160).ByteCount = 1

InSet(161).OpCode = "3E"
InSet(161).Nemo = "MVI A "
InSet(161).ByteCount = 2

InSet(162).OpCode = "06"
InSet(162).Nemo = "MVI B "
InSet(162).ByteCount = 2

InSet(163).OpCode = "0E"
InSet(163).Nemo = "MVI C "
InSet(163).ByteCount = 2

InSet(164).OpCode = "16"
InSet(164).Nemo = "MVI D "
InSet(164).ByteCount = 2

InSet(165).OpCode = "1E"
InSet(165).Nemo = "MVI E "
InSet(165).ByteCount = 2

InSet(166).OpCode = "26"
InSet(166).Nemo = "MVI H "
InSet(166).ByteCount = 2

InSet(167).OpCode = "2E"
InSet(167).Nemo = "MVI L "
InSet(167).ByteCount = 2

InSet(168).OpCode = "36"
InSet(168).Nemo = "MVI M "
InSet(168).ByteCount = 2

InSet(169).OpCode = "00"
InSet(169).Nemo = "NOP "
InSet(169).ByteCount = 1


InSet(170).OpCode = "B7"
InSet(170).Nemo = "ORA A"
InSet(170).ByteCount = 1

InSet(171).OpCode = "B0"
InSet(171).Nemo = "ORA B"
InSet(171).ByteCount = 1

InSet(172).OpCode = "B1"
InSet(172).Nemo = "ORA C"
InSet(172).ByteCount = 1

InSet(173).OpCode = "B2"
InSet(173).Nemo = "ORA D"
InSet(173).ByteCount = 1

InSet(174).OpCode = "B3"
InSet(174).Nemo = "ORA E"
InSet(174).ByteCount = 1

InSet(175).OpCode = "B4"
InSet(175).Nemo = "ORA H"
InSet(175).ByteCount = 1

InSet(176).OpCode = "B5"
InSet(176).Nemo = "ORA L"
InSet(176).ByteCount = 1

InSet(177).OpCode = "B6"
InSet(177).Nemo = "ORA M"
InSet(177).ByteCount = 1


InSet(178).OpCode = "F6"
InSet(178).Nemo = "ORI "
InSet(178).ByteCount = 2

InSet(179).OpCode = "D3"
InSet(179).Nemo = "OUT "
InSet(179).ByteCount = 2

InSet(180).OpCode = "E9"
InSet(180).Nemo = "PCHL "
InSet(180).ByteCount = 1

InSet(181).OpCode = "C1"
InSet(181).Nemo = "POP B"
InSet(181).ByteCount = 1

InSet(182).OpCode = "D1"
InSet(182).Nemo = "POP D"
InSet(182).ByteCount = 1

InSet(183).OpCode = "E1"
InSet(183).Nemo = "POP H"
InSet(183).ByteCount = 1

InSet(184).OpCode = "F1"
InSet(184).Nemo = "POP PSW"
InSet(184).ByteCount = 1

InSet(185).OpCode = "C5"
InSet(185).Nemo = "PUSH B"
InSet(185).ByteCount = 1

InSet(186).OpCode = "D5"
InSet(186).Nemo = "PUSH D"
InSet(186).ByteCount = 1

InSet(187).OpCode = "E5"
InSet(187).Nemo = "PUSH H"
InSet(187).ByteCount = 1

InSet(188).OpCode = "F5"
InSet(188).Nemo = "PUSH PSW"
InSet(188).ByteCount = 1

InSet(189).OpCode = "17"
InSet(189).Nemo = "RAL "
InSet(189).ByteCount = 1

InSet(190).OpCode = "1F"
InSet(190).Nemo = "RAR "
InSet(190).ByteCount = 1

InSet(191).OpCode = "D8"
InSet(191).Nemo = "RC "
InSet(191).ByteCount = 1



InSet(192).OpCode = "C9"
InSet(192).Nemo = "RET "
InSet(192).ByteCount = 1


InSet(193).OpCode = "20"
InSet(193).Nemo = "RIM "
InSet(193).ByteCount = 1

InSet(194).OpCode = "07"
InSet(194).Nemo = "RLC "
InSet(194).ByteCount = 1

InSet(195).OpCode = "F8"
InSet(195).Nemo = "RM "
InSet(195).ByteCount = 1

InSet(196).OpCode = "D0"
InSet(196).Nemo = "RNC "
InSet(196).ByteCount = 1

InSet(197).OpCode = "C0"
InSet(197).Nemo = "RNZ "
InSet(197).ByteCount = 1

InSet(198).OpCode = "F0"
InSet(198).Nemo = "RP "
InSet(198).ByteCount = 1

InSet(199).OpCode = "E8"
InSet(199).Nemo = "RPE "
InSet(199).ByteCount = 1

InSet(200).OpCode = "E0"
InSet(200).Nemo = "RPO "
InSet(200).ByteCount = 1

InSet(201).OpCode = "0F"
InSet(201).Nemo = "RRC "
InSet(201).ByteCount = 1

InSet(202).OpCode = "C7"
InSet(202).Nemo = "RST 0"
InSet(202).ByteCount = 1


InSet(203).OpCode = "CF"
InSet(203).Nemo = "RST 1"
InSet(203).ByteCount = 1

InSet(204).OpCode = "D7"
InSet(204).Nemo = "RST 2"
InSet(204).ByteCount = 1

InSet(205).OpCode = "DF"
InSet(205).Nemo = "RST 3"
InSet(205).ByteCount = 1

InSet(206).OpCode = "E7"
InSet(206).Nemo = "RST 4"
InSet(206).ByteCount = 1

InSet(207).OpCode = "EF"
InSet(207).Nemo = "RST 5"
InSet(207).ByteCount = 1

InSet(208).OpCode = "F7"
InSet(208).Nemo = "RST 6"
InSet(208).ByteCount = 1

InSet(209).OpCode = "FF"
InSet(209).Nemo = "RST 7"
InSet(209).ByteCount = 1

InSet(210).OpCode = "C8"
InSet(210).Nemo = "RZ "
InSet(210).ByteCount = 1

InSet(211).OpCode = "9F"
InSet(211).Nemo = "SBB A "
InSet(211).ByteCount = 1

InSet(212).OpCode = "98"
InSet(212).Nemo = "SBB B "
InSet(212).ByteCount = 1

InSet(213).OpCode = "99"
InSet(213).Nemo = "SBB C "
InSet(213).ByteCount = 1

InSet(214).OpCode = "9A"
InSet(214).Nemo = "SBB D "
InSet(214).ByteCount = 1

InSet(215).OpCode = "9B"
InSet(215).Nemo = "SBB E "
InSet(215).ByteCount = 1

InSet(216).OpCode = "9C"
InSet(216).Nemo = "SBB H "
InSet(216).ByteCount = 1

InSet(217).OpCode = "9D"
InSet(217).Nemo = "SBB L "
InSet(217).ByteCount = 1

InSet(218).OpCode = "9E"
InSet(218).Nemo = "SBB M "
InSet(218).ByteCount = 1

InSet(219).OpCode = "DE"
InSet(219).Nemo = "SBI "
InSet(219).ByteCount = 2

InSet(220).OpCode = "22"
InSet(220).Nemo = "SHLD "
InSet(220).ByteCount = 3








InSet(221).OpCode = "30"
InSet(221).Nemo = "SIM "
InSet(221).ByteCount = 1


InSet(222).OpCode = "F9"
InSet(222).Nemo = "SPHL "
InSet(222).ByteCount = 1

InSet(223).OpCode = "32"
InSet(223).Nemo = "STA "
InSet(223).ByteCount = 3

InSet(224).OpCode = "02"
InSet(224).Nemo = "STAX B"
InSet(224).ByteCount = 1

InSet(225).OpCode = "12"
InSet(225).Nemo = "STAX D"
InSet(225).ByteCount = 1

InSet(226).OpCode = "37"
InSet(226).Nemo = "STC "
InSet(226).ByteCount = 1

InSet(227).OpCode = "97"
InSet(227).Nemo = "SUB A "
InSet(227).ByteCount = 1

InSet(228).OpCode = "90"
InSet(228).Nemo = "SUB B "
InSet(228).ByteCount = 1

InSet(229).OpCode = "91"
InSet(229).Nemo = "SUB C "
InSet(229).ByteCount = 1

InSet(230).OpCode = "92"
InSet(230).Nemo = "SUB D "
InSet(230).ByteCount = 1

InSet(231).OpCode = "93"
InSet(231).Nemo = "SUB E "
InSet(231).ByteCount = 1

InSet(232).OpCode = "94"
InSet(232).Nemo = "SUB H "
InSet(232).ByteCount = 1

InSet(233).OpCode = "95"
InSet(233).Nemo = "SUB L "
InSet(233).ByteCount = 1

InSet(234).OpCode = "96"
InSet(234).Nemo = "SUB M "
InSet(234).ByteCount = 1

InSet(235).OpCode = "D6"
InSet(235).Nemo = "SUI "
InSet(235).ByteCount = 2

InSet(236).OpCode = "EB"
InSet(236).Nemo = "XCHG "
InSet(236).ByteCount = 1

InSet(237).OpCode = "AF"
InSet(237).Nemo = "XRA A"
InSet(237).ByteCount = 1


InSet(238).OpCode = "A8"
InSet(238).Nemo = "XRA B"
InSet(238).ByteCount = 1

InSet(239).OpCode = "A9"
InSet(239).Nemo = "XRA C"
InSet(239).ByteCount = 1

InSet(240).OpCode = "AA"
InSet(240).Nemo = "XRA D"
InSet(240).ByteCount = 1

InSet(241).OpCode = "AB"
InSet(241).Nemo = "XRA E"
InSet(241).ByteCount = 1

InSet(242).OpCode = "AC"
InSet(242).Nemo = "XRA H"
InSet(242).ByteCount = 1

InSet(243).OpCode = "AD"
InSet(243).Nemo = "XRA L"
InSet(243).ByteCount = 1

InSet(244).OpCode = "AE"
InSet(244).Nemo = "XRA M"
InSet(244).ByteCount = 1

InSet(245).OpCode = "EE"
InSet(245).Nemo = "XRI "
InSet(245).ByteCount = 2

InSet(246).OpCode = "E3"
InSet(246).Nemo = "XTML "
InSet(246).ByteCount = 1


End Sub

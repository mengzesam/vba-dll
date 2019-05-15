Attribute VB_Name = "WASP_TPRI"



'
' 宏1 Macro
' 江浩记录宏1998-4-5
'
   Private sgg, sff, saa As Double
   Private w1, w2, w3, w4, w5, w6, w7 As Double
   Private z0, z1, z2, z3, z4, z5, z6, z7, z8, z9, zt, zp
   Private pa, ta, pb, Mpa, tt, hg, hf, ha, sf, sg, sa, xa As Double
   Private vg, vf, va As Double
   Private zw1(23), zw2(11), zw3(7), zw4(7), zw5(5), zw6(5), zw7(5)
   Private zw8(7, 2), zw9(7, 2), zwz(2, 2), zws(2, 2) As Double
   Private it(0 To 1) As Double
   Private H(0 To 3) As Double
   Private H2(0 To 5, 0 To 6) As Double
   
Function swdat() '    *** 读水蒸气性质系数子程序 ***
 zw2(0) = 0.8438375405:    zw1(0) = 6824.687741: zw1(12) = -0.02616571843
 zw2(1) = 0.0005362162162: zw1(1) = -542.2063673: zw1(13) = 0.00152241179
 zw2(2) = 1.72:           zw1(2) = -20966.66205: zw1(14) = 0.02284279054
 zw2(3) = 0.07342278489:  zw1(3) = 39412.86787: zw1(15) = 242.1647003
 zw2(4) = 0.04975887:     zw1(4) = -67332.77739: zw1(16) = 1.269716088E-10
 zw2(5) = 0.65371543:     zw1(5) = 99023.81028: zw1(17) = 2.074838328E-07
 zw2(6) = 0.00000115:     zw1(6) = -109391.1774: zw1(18) = 2.17402035E-08
 zw2(7) = 0.000015108:    zw1(7) = 85908.41667: zw1(19) = 1.105710498E-09
 zw2(8) = 0.14188:        zw1(8) = -45111.68742: zw1(20) = 12.93441934
 zw2(9) = 7.002753165:    zw1(9) = 14181.38926: zw1(21) = 0.00001308119072
 zw2(10) = 0.0002995284926: zw1(10) = -2017.271113: zw1(22) = 6.047626338E-14
 zw2(11) = 0.204:         zw1(11) = 7.982692717: zw1(23) = 0

 zw3(0) = 523.5718623:    zw4(0) = 0.08565182058
 zw8(0, 0) = 0.06670375918: zw9(0, 0) = 13
 zw8(0, 1) = 1.388983801: zw9(0, 1) = 3
 zw8(0, 2) = 0:           zw9(0, 2) = 0
 
 zw3(1) = -2693.088365:   zw4(1) = -0.6547711697
 zw8(1, 0) = 0.08390104828: zw9(1, 0) = 18
 zw8(1, 1) = 0.02614670893: zw9(1, 1) = 2
 zw8(1, 2) = -0.03373439453: zw9(1, 2) = 1

 zw3(2) = 5745.984054:    zw4(2) = 0.4330662834
 zw8(2, 0) = 0.4520918904: zw9(2, 0) = 18
 zw8(2, 1) = 0.1069036614: zw9(2, 1) = 10
 zw8(2, 2) = 0:           zw9(2, 2) = 0

 zw3(3) = -6508.211677:    zw4(3) = -54.38923329
 zw8(3, 0) = -0.5975336707: zw9(3, 0) = 25
 zw8(3, 1) = -0.08847535804: zw9(3, 1) = 14
 zw8(3, 2) = 0:            zw9(3, 2) = 0
 
 zw3(4) = 4126.607219:    zw4(4) = 28.5606796
 zw8(4, 0) = 0.5958051609: zw9(4, 0) = 32
 zw8(4, 1) = -0.5159303373: zw9(4, 1) = 28
 zw8(4, 2) = 0.2075021122: zw9(4, 2) = 24

 zw3(5) = -1388.522425:   zw4(5) = 0
 zw8(5, 0) = 0.1190610271: zw9(5, 0) = 12
 zw8(5, 1) = -0.09867174132: zw9(5, 1) = 11
 zw8(5, 2) = 0:           zw9(5, 2) = 0

 zw3(6) = 193.6587558:    zw4(6) = 0
 zw8(6, 0) = 0.1683998803: zw9(6, 0) = 24
 zw8(6, 1) = -0.05809438001: zw9(6, 1) = 18
 zw8(6, 2) = 0:           zw9(6, 2) = 0

 zw3(7) = 0:              zw4(7) = 0
 zw8(7, 0) = 0.006552390126: zw9(7, 0) = 24
 zw8(7, 1) = 0.0005710218649: zw9(7, 1) = 14
 zw8(7, 2) = 0:           zw9(7, 2) = 0

 zwz(0, 0) = 0.4006073948: zws(0, 0) = 14:     zwz(0, 1) = 0
 zws(0, 1) = 0:        zwz(0, 2) = 0:          zws(0, 2) = 0

 zwz(1, 0) = 0.08636081627: zws(1, 0) = 19:    zwz(1, 1) = 0
 zws(1, 1) = 0:         zwz(1, 2) = 0:         zws(1, 2) = 0

 zwz(2, 0) = -0.8532322921: zws(2, 0) = 54:    zwz(2, 1) = 0.3460208816
 zws(2, 1) = 27:        zwz(2, 2) = 0:         zws(2, 2) = 0
 
 zw5(0) = 0:          zw6(0) = 1454.13:    zw7(0) = 15.74373327
 zw5(1) = -0.093996781: zw6(1) = -13790.159: zw7(1) = -34.17061978
 zw5(2) = 1.5549155:  zw6(2) = 52171.909:  zw7(2) = 19.31380707
 zw5(3) = -17.517633: zw6(3) = -98354.96:  zw7(3) = 0.763333333
 zw5(4) = -0.33890382: zw6(4) = 92218.23:  zw7(4) = 16.83599274
 zw5(5) = 0.027468338: zw6(5) = -33669.27: zw7(5) = 4.260321148
 
 H(0) = 1: H(1) = 0.978197: H(2) = 0.579829: H(3) = -0.202354
 
 H2(0, 0) = 0.5132047: H2(0, 1) = 0.2151778: H2(0, 2) = -0.2818107
 H2(0, 3) = 0.1778064: H2(0, 4) = -0.0417661: H2(0, 5) = 0: H2(0, 6) = 0
 H2(1, 0) = 0.3205656: H2(1, 1) = 0.7317883: H2(1, 2) = -1.070786
 H2(1, 3) = 0.460504: H2(1, 4) = 0: H2(1, 5) = -0.01578386: H2(1, 6) = 0
 H2(2, 0) = 0: H2(2, 1) = 1.241044: H2(2, 2) = -1.263184
 H2(2, 3) = 0.2340379: H2(2, 4) = 0: H2(2, 5) = 0: H2(2, 6) = 0
 H2(3, 0) = 0: H2(3, 1) = 1.476783: H2(3, 2) = 0: H2(3, 3) = -0.4924179
 H2(3, 4) = 0.1600435: H2(3, 5) = 0: H2(3, 6) = -0.003629481
 H2(4, 0) = -0.7782567: H2(4, 1) = 0: H2(4, 3) = 0: H2(4, 4) = 0
 H2(4, 5) = 0: H2(4, 6) = 0
 H2(5, 0) = 0.1885447: H2(5, 1) = 0: H2(5, 3) = 0: H2(5, 4) = 0
 H2(5, 5) = 0: H2(5, 6) = 0
 End Function


Function swtapb(ref) As Double
    swdat
     sswtapb (ref)
   swtapb = pb
End Function


Function swpata(ref) As Double
     swdat
     sswpata (ref)
    swpata = ta
End Function



Function swptah(ref1, ref2) As Double
    swdat
   Call sswpta(ref1, ref2)
   swptah = ha
End Function

Function swptas(ref1, ref2) As Double
     swdat
     Call sswpta(ref1, ref2)
   swptas = sa
End Function


Function swptav(ref1, ref2) As Double
    swdat
   Call sswpta(ref1, ref2)
   swptav = va
End Function


Function swptgh(ref1, ref2) As Double

   swdat
   ref1 = ref1 * 10.1972
  Call sswptg(ref1, ref2)
   swptgh = hg * 4.1868
End Function


Function swptgs(ref1, ref2) As Double
   swdat
   ref1 = ref1 * 10.1972
Call sswptg(ref1, ref2)
   swptgs = sg * 4.1868
End Function

Function swptgv(ref1, ref2) As Double
   swdat
   ref1 = ref1 * 10.1972
   Call sswptg(ref1, ref2)
   swptgv = vg
End Function

Function swptfh(ref1, ref2) As Double
   swdat
   ref1 = ref1 * 10.1972
   Call sswptf(ref1, ref2)
   swptfh = hf * 4.1868
End Function

Function swptfs(ref1, ref2) As Double
   swdat
   ref1 = ref1 * 10.1972
Call sswptf(ref1, ref2)
   swptfs = sf * 4.1868
End Function

Function swptfv(ref1, ref2) As Double
   swdat
   ref1 = ref1 * 10.1972
   Call sswptf(ref1, ref2)
   swptfv = vf
End Function


Function swpsah(ref1, ref2) As Double
    swdat
   Call sswpsa(ref1, ref2)
   swpsah = ha * 4.1868
End Function


Function swpsav(ref1, ref2) As Double
   swdat
  Call sswpsa(ref1, ref2)
   swpsav = va
End Function


Function swpsat(ref1, ref2) As Double
   swdat
  Call sswpsa(ref1, ref2)
   swpsat = ta
End Function

Function swpsax(ref1, ref2) As Double
    swdat
   Call sswpsa(ref1, ref2)
   swpsax = xa
End Function

Function swphas(ref1, ref2) As Double
    swdat
   Call sswpha(ref1, ref2)
   swphas = sa * 4.1868
End Function


Function swphav(ref1, ref2) As Double
    swdat
   Call sswpha(ref1, ref2)
   swphav = va
End Function


Function swphat(ref1, ref2) As Double
    swdat
   Call sswpha(ref1, ref2)
   swphat = ta
End Function


Function swphax(ref1, ref2) As Double
    swdat
   Call sswpha(ref1, ref2)
   swphax = xa
End Function

Function swhsav(ref1, ref2) As Double
    swdat
   Call sswhsa(ref1, ref2)
   swhsav = va
End Function


Function swhsap(ref1, ref2) As Double
   swdat
  Call sswhsa(ref1, ref2)
   swhsap = pa / 10.1972
End Function

Function swhsat(ref1, ref2) As Double
    swdat
   Call sswhsa(ref1, ref2)
   swhsat = ta
End Function

Function swhsax(ref1, ref2) As Double
    swdat
   Call sswhsa(ref1, ref2)
   swhsax = xa
End Function

Function swpxav(ref1, ref2) As Double
   swdat
  Call sswpxa(ref1, ref2)
   swpxav = va
End Function

Function swpxah(ref1, ref2) As Double
    swdat
   Call sswpxa(ref1, ref2)
   swpxah = ha * 4.1868
End Function

Function swpxas(ref1, ref2) As Double
    swdat
   Call sswpxa(ref1, ref2)
   swpxas = sa * 4.1868
End Function

Sub sswtapb(ta) '"B"---Subroutine PB=f(TA)
      zt = (ta + 273.15) / 647.3
     z1 = 1 - zt
     z2 = (((64.23285504 - 118.9646225 * z1) * z1 - 168.1706546) * z1 - 26.08023696) * z1
     z2 = (z2 - 7.691234564) * z1
     z3 = (20.9750676 * z1 + 4.16711732) * z1 + 1
     z4 = z1 / (1000000000# * z1 * z1 + 6)
     pb = 225.56 * Exp(z2 / z3 / zt - z4)
     pb = 0.0980665 * pb
     End Sub


     
Sub sswpata(pa)   '"A"---Subroutine TA=f(PA)
     pa = 10.1972 * pa
     ta = 100 * pa ^ 0.25
L11: sswtapb (ta)
     pb = 10.1972 * pb
     If Abs((pa - pb) / pa) > 0.000005 Then
       ta = ta + 25 * (pa - pb) / pb ^ 0.75
       GoTo L11
     Else
     End If
End Sub




Sub sswpta(pa, ta) '"F&G---Subroutine HA,SA,VA=f(PA,TA)
     sswtapb (ta)
     pa = 10.1972 * pa: pb = 10.1972 * pb
     If pa < pb Then
        sswptg pa, ta
        ha = 4.1868 * hg: sa = 4.1868 * sg: va = vg
     Else
        sswptf pa, ta
        ha = 4.1868 * hf: sa = 4.1868 * sf: va = vf
     End If
End Sub
 
Sub sswptf(pa, ta) '"F"---Subroutine HF,SF,VF=f(PA,TA)
     zp = pa / 225.56
     zt = (ta + 273.15) / 647.3
     z2 = zt ^ 11
     z3 = z2 * zt ^ 8
     z6 = zt * zt
     z0 = 1! - zw2(0) * z6 - zw2(1) / z6 ^ 3
     z9 = zw1(22) * zp ^ 4 / z3 / zt
     
'      Z1 = zw2(2) * Z0 * Z0 - 2! * zw2(3) * ZT + 2! * zw2(4) * ZP
 '     PRINT Z1
     z1 = z0 + Sqr(zw2(2) * z0 * z0 - 2! * zw2(3) * zt + 2! * zw2(4) * zp)
     z7 = z1 ^ (5 / 17)
     ZV = zw1(11) * zw2(4) / z7 + zw1(12) + zw1(13) * zt + zw1(14) * z6
     ZV = ZV + zw1(15) * (zw2(5) - zt) ^ 10 + zw1(16) / (zw2(6) + z3)
     ZV = ZV - (zw1(17) + (2! * zw1(18) + 3! * zw1(19) * zp) * zp) / (zw2(7) + z2)
     ZV = ZV - zw1(20) * z3 * (zw2(8) / zt + zt) * (zw2(10) - 3! / (zw2(9) + zp) ^ 4)
     ZV = ZV + 3! * zw1(21) * (zw2(11) - zt) * zp * zp + 4! * z9 / zp
    
'     Print ZV
    
     vf = 0.00317 * ZV
     z8 = ((zw1(19) * zp + zw1(18)) * zp + zw1(17)) * zp / (zw2(7) + z2) ^ 2
     z4 = 6! * zw2(1) / zt ^ 7 - 2! * zw2(0) * zt
     ZB = zw1(21) * zp ^ 3
     ZC = zw1(15) * (zw2(5) - zt) ^ 9
     ZH = 0
     For IZ = 9 To 0 Step -1
       ZH = ZH * zt + (IZ - 1) * zw1(IZ + 1)
     Next IZ
     ZH = zw1(11) * (z1 * (17! * (z1 / 29! - z0 / 12!) + 5! * zt * z4 / 12!) + zw2(3) * zt) / z7 - ZH
     z5 = zw1(12) - zw1(14) * z6 + ZC * (9 * zt + zw2(5))
     ZA = zw1(16) / (zw2(6) + z3) ^ 2
     z5 = z5 + (zw2(6) + 20 * z3) * ZA
     ZV = (1 / (zw2(9) + zp) ^ 3 + zw2(10) * zp) * zw1(20) * z3
     ZH = ZH + zw1(0) * zt + z5 * zp - (zw2(2) - 1) * zt * z0 * z4 / z7 * zw1(11) - z8 * (12 * z2 + zw2(7))
     ZH = ZH + (17 * zw2(8) + 19 * z6) * ZV / zt + zw2(11) * ZB + 21 * z9
     hf = ZH * 16.74796981
     ZS = 0
     For JZ = 9 To 1 Step -1
       ZS = ZS * zt + JZ * zw1(JZ + 1)
     Next JZ
     ZS = zw1(0) * Log(zt) - ZS + zw1(11) * ((5 * z1 / 12 - (zw2(2) - 1) * z0) * z4 + zw2(3)) / z7
     ZS = ZS + zp * (10 * ZC - zw1(13) - 2 * zw1(14) * zt + 19 * ZA * z3 / zt) - 11 * z2 * z8 / zt
     ZS = ZS + (18 * zw2(8) / z6 + 20) * ZV + ZB + 20 * z9 / zt
     sf = ZS * 0.025873582
    End Sub


Sub sswptg(pa, ta) '"G"---Subroutine HG,SG,VG=f(PA,TA)
     zp = pa / 225.56
     zt = (ta + 273.15) / 647.3
     z0 = zw7(0) + (zw7(1) + zw7(2) * zt) * zt
     ZZ = zp * (zp / z0) ^ 10
     z1 = zw7(1) + 2 * zw7(2) * zt
     z2 = Exp(zw7(3) * (1 - zt))
     z3 = 0: z4 = 0: z5 = 0
     For IZ = 0 To 4
       z6 = 0: z7 = 0
       For JZ = 0 To 2
     z8 = z2 ^ zw9(IZ, JZ)
     z6 = z6 + zw8(IZ, JZ) * z8
     z7 = z7 + zw8(IZ, JZ) * (1 + zw7(3) * zw9(IZ, JZ) * zt) * z8
     z3 = z3 + zw7(3) * zp ^ (IZ + 1) * zw9(IZ, JZ) * zw8(IZ, JZ) * z8
       Next JZ
       z4 = z4 + z6 * (IZ + 1) * zp ^ IZ
       z5 = z5 + z7 * zp ^ (IZ + 1)
     Next IZ
     z8 = 0
     For IZ = 0 To 2
       z6 = 0: z9 = 0
       For JZ = 0 To 2
     z6 = z6 + zw8(IZ + 5, JZ) * z2 ^ zw9(IZ + 5, JZ)
     z9 = z9 + zwz(IZ, JZ) * z2 ^ zws(IZ, JZ)
       Next JZ
       ZS = zp ^ (IZ + 3)
       z8 = z8 + z6 * (IZ + 4) * ZS / (1 + z9 * zp * ZS) ^ 2
     Next IZ
     z6 = 0
     If zp < 0.1 Then GoTo L13
     For IZ = 0 To 6
       z6 = z6 * z2 + zw3(IZ)
     Next IZ
     z6 = 11 * z6 * (zp / z0) ^ 10
L13: vg = 0.00317 * (zw7(5) * zt / zp - z4 - z8 + z6)
     z4 = 0
     For IZ = 0 To 4
       z4 = z4 * zt + (3 - IZ) * zw4(IZ)
     Next IZ
     z6 = 0
     If zp < 0.005 Then GoTo L14
     For IZ = 0 To 2
       z8 = 0
       For JZ = 0 To 2
     z7 = 0: z9 = 0
     For KZ = 0 To 2
       z7 = z7 + zws(IZ, KZ) * zwz(IZ, KZ) * z2 ^ zws(IZ, KZ)
       z9 = z9 + zwz(IZ, KZ) * z2 ^ zws(IZ, JZ)
     Next KZ
     ZS = zp ^ (IZ + 3)
     ZV = z2 ^ zw9(IZ + 5, JZ) * (1 + (zw9(IZ + 5, JZ) - z7 / (z9 + 1 / ZS)) * zw7(3) * zt)
     z8 = z8 + ZV * zw8(IZ + 5, JZ)
       Next JZ
       z6 = z6 + z8 / (z9 + 1 / zp / ZS)
     Next IZ
L14: z8 = 0
     If zp < 0.1 Then GoTo L15
     For IZ = 0 To 6
       z8 = z8 * z2 + zw3(IZ) * (1 + zt * (10 + z1 / z0 + zw7(3) * (6 - IZ)))
     Next IZ
L15: ZH = zw7(4) * zt - z4 - z5 - z6 + z8 * ZZ
     hg = ZH * 16.74796981
     z4 = 0
     For IZ = 0 To 4
       z4 = z4 * zt + (4 - IZ) * zw4(IZ)
     Next IZ
     z5 = 0
     If zp < 0.005 Then GoTo L16
     For IZ = 0 To 2
       z6 = 0
       For JZ = 0 To 2
     z7 = 0: z8 = 0
     For KZ = 0 To 2
       z9 = z2 ^ zws(IZ, KZ) * zwz(IZ, KZ)
       z7 = z7 + z9 * zws(IZ, KZ)
       z8 = z8 + z9
     Next KZ
     ZS = 1 / zp ^ (IZ + 4) + z8
     z6 = z6 + (zw9(IZ + 5, JZ) - z7 / ZS) * zw8(IZ + 5, JZ) * z2 ^ zw9(IZ + 5, JZ)
       Next JZ
       z5 = z5 + z6 / ZS
     Next IZ
L16: z5 = z5 * zw7(3)
     z6 = 0
     If zp < 0.1 Then GoTo L17
     For i% = 0 To 6
       z6 = z6 * z2 + (10! * z1 / z0 + zw7(3) * (6 - i%)) * zw3(i%)
     Next
     z6 = z6 * ZZ
L17: ZS = zw7(4) * Log(zt) - zw7(5) * Log(zp) - z4 / zt - z3 - z5 + z6
     sg = ZS * 0.025873582
     End Sub



Sub sswpsa(pa, sa) '"PS"---Subroutine HA,TA,VA,XA=f(PA,SA)
     sa = sa / 4.1868
     sswpata (pa)
     pa = 10.1972 * pa
     sswptf pa, ta
     sswptg pa, ta
   '  sf = sf *4.186868: sg = sg / 4.1868
L21: If sa > sf Then GoTo L18
     xa = 0: ta = 376 * sa
L19: sswptf pa, ta
     If Abs(sa - sf) < 0.000005 Then Let va = vf: ha = hf: Exit Sub
     ta = ta + 376 * (sa - sf)
     GoTo L19

L18: If sa > sg Then GoTo L20
     xa = (sa - sf) / (sg - sf)
     va = vf + xa * (vg - vf)
     ha = hf + xa * (hg - hf)
     Exit Sub

L20: xa = 1: w3 = 1: w1 = pa / 1000: zt = (ta + 273.15) / 1000: w2 = sa
      w4 = 1.44 + 0.3746 * Log(zt) - 0.1102 * Log(w1)
     If w4 < w2 Then GoTo L24
     w2 = w4
    GoTo L22
L24:  zt = (w2 - 1.44) * 1.6696 + 0.2943 * Log(w1)
     zt = Exp(zt)
     ta = 1000 * zt - 273.15
     sswptg pa, ta
   '  sg = sg / 4.1868
L22: If Abs(sa - sg) < 0.000001 Then Let va = vg: ha = hg: Exit Sub
     If w3 > 1 Then Let w5 = w4 + (w2 - w4) * (sa - w6) / (sg - w6): GoTo L23
     w5 = w2 + 0.8 * (sa - sg)
L23: w3 = w3 + 1: w6 = sg: w4 = w2: w2 = w5
     GoTo L24
End Sub



Sub sswpha(pa, ha) '"PH"---Subroutine SA,TA,VA,XA=f(PA,HA)
     sswpata (pa)
     pa = pa * 10.1972
     sswptf pa, ta: sswptg pa, ta
     ha = ha / 4.1868
     If ha > hf Then GoTo L25
     xa = 0: ta = 0.9189 * ha
L26:  sswptf pa, ta
     If Abs(ha - hf) < 0.001 Then
       va = vf: sa = sf: Exit Sub
     End If
     ta = ta + ha - hf
     GoTo L26
L25: If ha > hg Then GoTo L27
     xa = (ha - hf) / (hg - hf)
     va = vf + xa * (vg - vf)
     sa = sf + xa * (sg - sf)
     Exit Sub
L27: xa = 1: w3 = 1: w1 = ha / 1000: zt = (ta + 273.15) / 1000
     w2 = 0.08218 + (0.1873 + 0.085 * zt) ^ 2 / 0.085
     If w2 < w1 Then GoTo L28
     w1 = w2
     GoTo L29
L28: zt = Sqr(w1 / 0.085 - 0.96685) - 2.2035
     ta = 1000 * zt - 273.15
     sswptg pa, ta
L29: If Abs(ha - hg) < 0.005 Then Let va = vg: sa = sg: Exit Sub
     If w3 > 1 Then GoTo L30
     w4 = w1 + 0.0002 * (ha - hg)
     GoTo L31
L30: w4 = w2 + (w1 - w2) * (ha - w5) / (hg - w5)
L31: w3 = w3 + 1: w2 = w1: w5 = hg: w1 = w4
     GoTo L28
End Sub



Sub sswhsa(ha, sa)  '"HS"---Subroutine PA,TA,VA,XA=f(HA,SA)
     ha = ha / 4.1868
     sa = sa / 4.1868
     hg = 0
     For IZ = 0 To 5
       hg = hg * sa + zw6(IZ)
     Next IZ
     If ha > hg Then GoTo L32
     w1 = (0.956 + 1 / sa) * ha + (258 - 0.502 * ha) * sa - 763
L34: w2 = 0
     For IZ = 1 To 5
       w2 = w2 * w1 / 100 + zw5(IZ)
     Next IZ
     ta = (ha - w2) / sa - 273.15
     If Abs(ta - w1) < 0.005 Then GoTo L33
     w1 = ta
     GoTo L34
L33: swtapb (ta)
     pa = pb * 10.1972
     sswptg pa, ta
     sswptf pa, ta
     xa = (ha - hf) / (hg - hf)
     va = vf + xa * (vg - vf)
     Exit Sub
L32: xa = 1: w3 = 1: w4 = ha / 1000: w5 = sa
L40: zt = Sqr(w4 / 0.085 - 0.96685) - 2.2035
     zp = Exp(9.071999 * (1.44 - w5) + 3.4 * Log(zt))
     ta = 1000 * zt - 273.15: pa = 1000 * zp
     If w3 > 1 Then GoTo L35
     w2 = ta
     sswpata (pa / 10.1972)
     If w2 >= ta Then GoTo L36
     zt = (ta + 273.15) / 1000
     w4 = (0.085 * zt + 0.1873) ^ 2 / 0.085 + 0.0822
     w5 = 1.44 + 0.3746 * Log(zt) - 0.1102 * Log(zp)
L36: ta = w2
L35: sswptg pa, ta
     If Abs(ha - hg) < 0.05 And Abs(sa - sg) < 0.00001 Then GoTo L37
     If w3 > 1 Then GoTo L38
     WW = w4 + 0.0008 * (ha - hg)
     W9 = w5 + 0.7 * (sa - sg)
     GoTo L39
L38: WW = w6 + (w4 - w6) * (ha - W8) / (hg - W8)
     W9 = w1 + (w5 - w1) * (sa - w7) / (sg - w7)
L39: w3 = w3 + 1: w7 = sg: w1 = w5: w5 = W9
     W8 = hg: w6 = w4: w4 = WW
     GoTo L40
L37: va = vg: Exit Sub
End Sub


Sub sswpxa(pa, xa) '"PX"---Subroutine HA,SA,VA,TA=f(PA,XA)
     swpata (pa)
     pa = pa * 10.1972
     sswptf pa, ta
     sswptg pa, ta
     va = vf + xa * (vg - vf)
     ha = hf + xa * (hg - hf)
     sa = sf + xa * (sg - sf)
End Sub

Function swptgu(pa, ta) As Double
  T = (ta + 237.15) / 647.27: P = pa / 22.115
  d = 1 / swptav(pa, ta) / 317.763
  
  u0 = Sqr(T) / (H(0) / T ^ 0 + H(1) / T ^ 1 + H(2) / T ^ 2 + H(3) / T ^ 3)
  
  u1 = 0
  For i = 0 To 5 Step 1
    For j = 0 To 6 Step 1
     u1 = u1 + H2(i, j) * ((1 / T - 1) ^ i * (d - 1) ^ j)
    Next j
  Next i
  
  u1 = Exp(u1 * d)
     
  U = u0 * u1 * 0.000055071
  swptgu = U
End Function

Function swptfu(pa, ta)
 U = (-2.513489692E-08 * ta ^ 5 + 0.00001840513152 * ta ^ 4 _
     - 0.005141450708 * ta ^ 3 + 0.6959923173 * ta ^ 2 - 48.66015769 * ta _
     + 1758.965976) * 10 ^ -6
 swptfu = U
End Function







import csv
import json
import os
import sys
import subprocess
import sqlite3
import hashlib
import zlib
import base64
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime

APP_TITLE = 'Cotyland | Comparador de Precios y Etiquetas'
APP_VERSION = '7.5.6'
# Equivalencias verificadas y literales extraidas del Libro1.xlsx provisto por Cotyland.
# Se conservan los codigos como texto, sin quitar/agregar ceros y sin inferir SKU.
_BARCODE_DATA_B85 = (
    'c-ox3OOxAJmL~XDU}@84u9$Qe?i0Z5KoCxcL=sGpbh)c@%ton*j?!dAC`}5LvFabpVis+^ncvW6GuFM$W#%87^SCd7pu4NevLcz`I3F(V<J|K;|M@@v'
    '!B)9YrO9)fm-!$5_y_nOxfP~_pZoTZnQXQ07yUllw1;}#y|0(}`w{-<x?etK*V)74^`As(Ga&$@`osVJM}oxvlhWD>afDO=Zz}j%FNSWtuD1uk8M0kH'
    ')E-}({MTQy>?E=)A#!chE5a&@%Bn&OQK%dNL@tZGvJ$bDckRBr@3RE}HhosF7j?F4SM=Wxnax)>55j}w8i!=6LNkcUr7lvC`+nW6`c0jw{NqQWs*I^b'
    'najK?3i=F!a8&)Ug%|gIc2}?0{kyi^PHC)&G&<5W^}1fv@9PQSfhNhQFkk%B6ZPvAOj?`mI{07P`o3GX|NO7`@3-0d_}<OB5```e(5k{MWMT3V%Fu7?'
    'ENkw%rY73`VW{`}Y`3ltZN02<_V>qYTzy8=qAGO(D_crN6tFk+8RC<GtH)~u@Q7t$rG!~2OFu>_t24w+x9w8Eq2I$!>DSBG@E8EU4sSqsk(b#iJGzO4'
    '#g)n!x6baTgKziEaoK09q24artZwRMyXopcC@Y&wCiFjN2{*8bVVFbL@7;Hm5=h`brLcYlbFCu%1|HtwN5pfKsU^(zaOjo}1?XwJz*#BWQ53?a!-qY3'
    '5A<ON@TAwfb@!)ZyT{qy)ky8(Oa(_#Nc=YpCv#;Z+^*gMR^P6(eZPfaTyOgOZqv2fL!Vvk+F@w>R|tB$Lk#z&R_28&0Lb_;mMS+XC>-iTKfo{-{kn#k'
    '+qG|SrC<}onOCrf+xPX}?KOnDRvugC7Ip-T(lE3XLRGWS?f$lfoslW&@I2=!OE`WO{;AyPDlN+=A!wAasPNxibJu2K0vSG7$oRo}7V@@*rOle-X18u@'
    'IG}s?(W039Y7xE~wo4(T$la!`fP3W03%V_iE2{_Ka^2AFY1Lae8ZGeB>@nLO7rgq2axHR)D+-C+UBDS|1J`N>_k7SRTzdGZx7n`WWv3zzSrYW_wyEDA'
    '*PfFKqt8HhFqyaYHoJ#|yzWQu$dES6+5>J{ya@Km`2%k6&maGMIjlRPaG4z1p$CQ_0<Jc2ju_Yfd^zr>Yu>hxqcNH|#_z-0)$ic{4@18W<Acc@`Zv9s'
    'OaZ6=yAX*s9a|a4R(=-}@_=SEZAD&qIREqIa7RFcXFDF<?!6Y;Xb)z}@B7O4?Y8HOa<{uRlJ8bnOk3*wIiL+pQrmQkj*mqFG639XFobXg<&{pk%&{fr'
    'H5}j6bY0)$S#Vcv=9nx{!{1UwxUx%PR)koNHU=a5tABqErY2y(atalPXU>Ng#D|PTDiC1kI?6Wv5=1k97XZo`vWoISUghV=yZW)-1EV=ihY5rp3?wQk'
    '3lueAaE@t#+?eozL%)L?WW0*Yj@jkm5$`@A%}(aq4P5x}2V`ZQZ&{_S25`5^aJ%yQz_Er4_!e)|?Xh0dxK{MwLP$TZQo}85a_0Itu0jywu5XFE*V}GW'
    'uYp|x#|P%QZ`b#*GB^DVfO&{2Hv;&kskED?vSPl;5DT~=?_0bu_Zf;R{__??-ED{dkX`-$<-;!@e);Lk=Zt@Splfus$?#?yy#%&iSY}0H=FCfm_P$-$'
    'O$W;Xqh7UJ<Zs&+<Rtp9)`9C{f(|mh)t>EXouD`U2F@qoH8rgd-6y+lRXeCkm;hBO4RVRmo-^Bm#X$#!E8KKTkfiV`Z3qG8b4;C=%JcP7@pW-d-PFT='
    'p4#C1tgO^LA+BRLz+}Phg;UVd&KU0pX>z~mL58*SI3T30fc|i{A$S-i%xO?T$_r!Jw8OgF&d-7o+DVtX&_yv10K(jc|1pnIYU9^I=NHxic6Hm(et|Cn'
    'aqiN9RK=1Z?hh+3%*y_V2l)>79I(PI2nYL1^HaE?6$*wqMhi?}u|Bqs)6qy#0w=Fnc!o^^JK0{KUk>#OK5PN}FMIEBZiE#=jM2){k4C~oT^h~p(8Jw='
    '7-w5XST8q$AYEJnxa&8I0ock=qc{%38GRQ6$-+sFz?EnV%>ljb4u|=9sjODvOLTF89#p1n*Pd+}5dSnxr}AO}sC1VQcI^)I0XR%eJ1pHP$bdkI32Ou~'
    'llUQSLIXF!1-zkOtosMgQEX&RZdzb6{MCW0Hw(f7>D%Mb1D_3Mhq=9s%7WW{a5R$BxvCe7$NvKueh)ilj<Tv8lZ2xs6DIkt+ux0LI%-rUnKlX)^fR7L'
    'n6-M*zJXbD*b-nXA+!~4ri#L_*<uX>Kn=%e&4vpWE!-uJ<6AF)h)@h%7X}y?C_(<)1TZ1xQJQ+QtBLD5@<(7C#|>@6cR=IreFaLl0+u7f1c5A81rT`q'
    ';RsW+ZQ*i-`yMtDDyo1Xl*;QkR}yq-ttw_=sNj}iB&bZ|Cv}J3sXT0ynGox|jKumfF%G4VTdt%@ccaJaUYN_YK_K>|__5ZJYEh(PzC*nfxZ~S)SbzXB'
    'B{O9t6A4*R=Fu#%A5p0iRy|O$$c$8QudI6faqR@4J!p*WKmI>5m2}(V#%(YGD7h}JvW_vEQrIG(kokevj8>ZN%jL30g6!&FKYaQ0`D^y^^Dm#j{Pgve'
    '-$QVh!-XYT0Kh#|lr~I$_!ioQJ%1B`F5vlVIf5vVc?1Ykok84&g?9*|9iN-h@=D;rsmkIEfc9ak-0opZdjnSm!-nGk>RUlaKu8H&LG!{aj++cbnq>!K'
    'A-oPIBkPB5MdES2@7CF?(K3NJVEjIouxf<SwBD4c@(1`&lqvC^q#alVomcS;HvJvwa<C_WXSVR}eT{0*I(x|AAa{qji#>$5FCm0_2Eyd|$B#Flk$t-f'
    '2MO-%3OF!*HqQWyQ|sIM4a|O1?|>hmoVZ4b4uwYlQ74`M{AG$Gi-=>M!J$(*hJk^gW!JLh0+<ubQJ%u|x9b-7MBls(_S_s52jMS!gSV?nXjaGi5k&NL'
    'yLE@d&7qZORVz$oa_hy6lFeC;a?`Eff<WFM_xm=x((>moa2Q^Hdlig>#EC!+0+CXv@OF1f83T3;@>8C&hpfoq(Cu%=)LOb#1ZrJYXVg4V7m0NqKm7dF'
    'hoNIe2J4>tslErLW#FTOyC^RbGWr`@KY+gsX!u6e+ly{wUL*)XrbQTrDJ6^Y43$3xmCs}0ezqE~IO*3y!hz<s*tA=a(OB}l9geI<Hg`RO9#B{L>(3c#'
    'K6zrFfm^cjDn0+K*JSH<3jya9r*cG4M1mp`kj~oT^*uu)g5AIg989xEBQOg8;}<RN%4ox9v-kQv-WWsw2<P-_3qs%a4YIvs|0<)Wk*rx4Ci)B9RYi2~'
    'Kz`1$E99`RGJY(GX$3kta^o2cFj#sajH*2xx_x#W{&ei#qhkYR4u&{pXExrBO7NN3w1)vs5vt=gTtC|tn#JP#Bye;f^MWO4(IlK~Nsb8O#1iKP9S@+^'
    '9=hxV(CK@PQiW&edC4*|>es;c(3S)z83ak-BAdFY7jPtyn=~zK5Eu$D6p$iM@HHn<e*86D6qdHB^18S1hXiQgX2+r74IHiL&h9~H2CbkryWJtZHzpWs'
    'K~E~BH&?^$%KY8|!p_|7koVrVFuhm6eaS)x>JJ;gH=9acOOuQbu0vJjaeQ>A1vn5?8_z2r$aW47U*|FcNHp<RBvBB6<hvvQ4E6hd>wPF_)p~$wBPb;E'
    'Q1BHRc76^E9z0{Bk}`p6;V{r}_OwU+5r(gl+2flWZ}YM2QdVA3shsD?A>DQ@@+MR(d2Qhkw7}X)x?i_aJH}v3GG^mV5Tqb&)A^x+4E5sz=6UM<u<ZWO'
    '#&|<JE4!Zdt{w&uabYXA1A9Ii*r+(&C<WajB|$yDDRNPGKTOGjhQo03Er6kIaNRDm%stNg#Z(eC9N>*20@x=3P8rQ-0HH-j&r4Yb8(yWWJOPQ81{`^}'
    '&($?-DCQn2@ir<uCVI@?)Z2D^igghbU9CjAwKrY6!DF{?A3CqYI7cHp^x{uX0=sU1z!~-D$*&PFTyty(P&nX!!sT$xu27&UV3mS&m7}Tx%u)t7SjiIF'
    '6rbVTUVlii<P=NM4ePMvMXcZtd5zl(56&E{=P)2Q=J5cIX3cHRu(u`*hR?KiUKUdH$`V<(a+eE^x;)2Qcl()HC4CQ7I_C<)8x{6Z5v`YHRh)xu$QR_i'
    'k+_9cNK=oQ6lg7F_(Z8^P>_p8M&rLMu$ZU{N##W@&?07#m`9=N<rsn--lU4ol+OfDfYwEI3GTk@*DW6rM7;0L5YaMRTF`CanzdesFJ;OzQhpxTxrpqX'
    'ViiPHRQ?((L`wch%TM<9g!e2PpzyYEg9;`r6QFqQCb!BYCwCv2I(^u>zi*vQ9cadSt}V4!R&$oq0HnLE-;?_Q&8r~bef#axj~{;g>FfCYCagCoXnI82'
    'Z`(I0w`>;uKD9SE|AS*(2~xNS07)lED-g(GHr9BBzd5#R@_cUVeZLM^)%bb>ZOvoBr7p)<?(Bw>>~5xFvH?L>E`sGinlb^|^jj2c_Sx5OUw;4e<LC4F'
    'InZYm(yxvC=>c}Ni=;s7O+TV2BSm>e@unN<+j{hT6`!xpo?o{;Y~N)+3ir@4+#{Xk`*zzlAUNUnMdR%<-rXz^zh?P@m>Zm?l9^k*TzNca;MpFR9BJbx'
    'IELJ?1LOc#_;q&o)<Ma_yz(cjC_%N9$b+0SKCf6~<#2;+8riVw&=9)p7wg1_?yntSZ74QL0Xkb54DtNY4le+?y@~rBgl4r&JnZ|-1?<$kRo)4>b!8HR'
    'q0l~wqcjD3;HKax$82^KkMJzM!v#$hy=HJD7crzK1t<Dl>&D!$A!Ka&f>Thq4J;&EOo4f0I0y>Jcr@RTC!cI^TAiwsb~16`nuXQgWT7S7quA*D1MHC%'
    'M#cP)*;AoUGCwDEiUpC%gcMj~zXihL1S@iqCK1F*7E@IjIHTg_XUdUd6u$jPAqppjQYsK7jzR-7CX8`W2q2}v4Ov7WKLZK3FGOi}*Zb(UL}5xQZ*<L_'
    '#nu7voVgvII!5tctdTAm4Buzklwm#$IWXxyeT0Pr)e*kwxWlt{_2J9M&zUiQglYL>_FsPa>EpkA$Ta*9cSfuQ(Gb2)u`!eGGxkl_!-l`Cy@wk%4(Wi}'
    '3+hwL78n=YNEFW0<3`rkfOB1rnJ370!Q$q8kozM_;`<|gHQ(UP>mAxZQJ2a#FqYTrJCGt|c7+C>T)moX2IGR@a5L~B0Znm-!+;SYM+dgfQ@j55@7b5n'
    '|L2EHXme8fB&v_U-ekK^zkK+befsU=FW>(8|Ni)axomw8cd7D0SCUuw9tE;3oekKY`-$obTS}wyn&%<pioN;ok4+7iZ2PW-6&Y&sWRay7jb&t7d%Fs1'
    'J$eV-?E9Z10Nx&Gb%-Ru9i4pTe~w6o<DyGQV5mYCXfi5fSZC*;VzLf6%6eBL=X={8-{Lg(Al0{M#4%{ZxvAfGbu^kO6hUpU1LcOiLW_E}f-Ug|J+lM)'
    'xZmU4fsDUdPf4BA-N3=3<2)ffuHhaTj)x1>O2rqeGiq<v7?HwS86ophIID|P{Bu#S$tAt^TpPaq7Iph2FxnRX#kM_y%7}_md*2;0eWU;Svp0!h*r~*s'
    'U_%Q@=90gMtUkb5_aLyez-Rz2v-!!WlP<HXJ2I8QfZqDeUBDcYpFr9wV3EC+u~t3%w2V9{V6ORS1Oo0%8Ao6P6-QIR#sewv2#gs{fsIBGCWfs+8m=Q2'
    'fq1uWSF|T|egfvI#xjl3w)9tmvC@qcb~#5lW<M;t;c&OCA2YaO+U*aKxJrd8&|Rzzw0Lx5z)QCM$A4!@w**FoLq)$V%Rv_Qf)0c*$BuvQv+U1*K+V3x'
    'Km7R*?#yPO;ebL$U@Z*!-HH8L3a<$Y9LC2{$P^eQf)Zx8ukqx)>xZ|<nj&eSE2?Md4AU(Z@aG3M;b^p<NR^O!5`q}IJ@Q06cnPUUNcD51Zm<*rC6xAO'
    '-u?8XaPSWpk-^I)Z^Uc)*PpX1C19~%MLT-|Yn$ic*cehTCK$UKL*ZIjufMY-4u9W@7>VlG!yL^CiBR4GZeM0_HKBUsjO~P?OoUO-!KfRPT}e$i-e^uJ'
    '35a<P1R=^1L^2U&oCy0gP&p;7B584%6!+_7N+v*IVwDlq3xJYInGVTDB>gOsV;f~9L;-Iuy-qF5Qq4^&+<H9*{SI`sXH1bm&_&=3+EQtG86reW?}m2M'
    '9k(zs)lC66`i-Gw61np*@s_ti+jl=Fu)tjo^>%ecCxP4k9vD9}17@cSCrKhHoLnDBCFX{i;%!y7zK3%k_TEyY)cG^i|M<Je@a#|bo$<j-C`dhvAxV~+'
    'LYQEiSMpg1O%QHc@llh0hJKdb#5q`kW>ffgTe2-ZLCf2Pze1D`Qr_?tu-L(Ni7@h+nZzwqJ0O!O^`IkD)Uxsn!T)>)sXf~xS8kK=ky<n$hkG+7NYzTr'
    'VJ`dRexwh)LS4r}ED<7vtLoxK5OA`Uilf+7rH_=sFDssfNWKX(SHcJRgJj}CGT#GM!Xdix+8D>;O6gCf`*{H@ZU8ums@cQ|E@hlxbr~Qm7;WP#m8TyC'
    '>aYoMnbr7W!blQKDR=~BZjavIO#^DA$ZxJPX+XIIl^WU2CcPI_S=dTz<vmeKuzU9@P_%F4#1p>(v;yfW_s|{|Ru3%$E=d9HsMd^je7|cC-PW5>4a~v4'
    'b972rM2dcqE-xXnTOUCI96KK!!{e^0-vTXWH|!~5GfpL)B4FJ&-ku=PCjwi%@L3t4NTxe&i57(&+EYCW+863>xE&rrrGPzu0>=G{^}XnyeS3syxkdpe'
    'njWkwFbp&t)<UNXw(15@|A#Kop7CG7>vrUPKkmKIW=8VfXte!3+fm@5xL97SO5{Eh<vwd)Om7k{T$OOFwLhPgDrCyZQ<mes#`#H>u$XUe?rFD3P#6}A'
    'i6E`L(NyMa;32?)mVv}+N|RY<;zsAc{{9u!Py8?3AUVcNGJjVXo#KzDh+7SCZb1>GkkA;5*$h2QcXnmVQ<c9O$JXh>hsB{@5C;UC6(lQ;hr2yCdssCx'
    'Wp&#X4n`?mzLqvyZN_khqt4vwS0lK8{CC7d)4|{)fKtdDF|SRE>4K|O6@#nu5$0@o^*HFxe`Xa@@k%RFj77b8%=UNvyB1RfQH|vBoqeuS=bu~n$V3e?'
    '(P(ICL^YdFM;}ZdYsIyqc*XGtTUmN{FGzZ7gxW7KK&D5Hc@kVmq*_^j?*p$Jlb*+mUylODr;~<m7@vb=E2UGgrry9lJGhXzH0pQyF=$BiuJFRxF~~qU'
    '6#f1+1^_MUVSe#Up}fmkaYS8&Nar+6c?c_`b+h8Nm|3}CM61Mj<wlsJcym7I@b_Eb=57gP{AMAvkYY6i|Dm-{;!bn75EgJ(m~(Z-V6P_e`<xJplWEvx'
    '{8Ds2!&5Zjy%ZP;_#J58^`=0vzdqvv<N=Lo2A3D?xl99ig9av?`4^y_7~@JY3atpfFKyBncs?38CA*3QMy`#wcP@NRAsm$v>sxljkTJ6bMzbNSX>hAi'
    '^4qtY)doJ_l-FVstz2rI$HXK3JS<Ny{w=VYsyN@UxmS~vWXT}F0(WWmZiL$zzGc!w$zlC6T+$e;I|AD_^_I-L=|SZ58{ode=L~8&+J2WZZ$L&6n0X`4'
    'GH+sdcaaB4QRJ6V_w{W(bUwhl$n_-qM0q!_)|@>y3sY}lp28fx9qQXdIz6J&n2F=eR+zX|*o#nf793Z2bjJ1Rz2no0;NCIi(@4NnkGB{?6>nI5cD6u#'
    '^Gq3c#WRTgX2A|UceAZKO5wb3ho$q9gJ>zOuxQ%#F)Ttz`7|!MQOWMh!Gl*gtXxiHi_!)mMdW4m0!a8@+YxX$5Fi<x&xB8A!kN%d!-GuKrs0k>1A&Ff'
    '>TI56jJbjvnb1G{_Tx`qKK+*6{`AYIf6tbmK4hB@zkbM?&+xA(270pg+UMWM(&jU3k4Ju9)DYu!C%1**ENu)SR>jQMZds*Z-aK03xE?-e1lMB>!4X%q'
    '2>yU6JrPT)5G7|PvITk|o9+;!(|uA*RmG&3@i_{aFylxe&OM1g0|EJaaVMXzEoJzULm3pjSQA2E9F#Lc;I7V<_#QwZK%c2bV>5AV=7msAyGGXp4X+R`'
    'u#fRHUqA8ZOYm65yK->Vt8{!O#E=>D(odz`;?UG+a5fIkz8GvD$5y@?!RZL6zXuM>Ft~65kl|mAmshf1%q{Or_Y0r>UERC|#szA3RlFKskJ@<qfCe5j'
    'I=Rug+GF0g595Qqj)Ra=`ukDo3Egh@m<1<GbF7!>!D=(Lp{sU{mLLS9oICh9DMPTU&pWsTb{JZ(3Q@2e;|cZpDD`zwI0ZdNXMFwMOFV@hEyEpfMa*`k'
    '0Jr%AfuS<SxJcKWugXH^WeBCVVsveEn;ix@%pZ#w6y)wt^28IGkL<)qIoRTXXTe2Pz7!w$JZb0N<?SL^@Fi~+hzl%#n!@X&p(|UAW!7$cXt(>C0*SL('
    'wM$4XLeAZI;V&oJVVr^Ij&giIz$Lw6x9HMkqbHm)W|Eb@ghT1#ECRr$-rwOx4_7`0kIrVW<uZ@?r6c|`lT~U<bSuFUsK`+<U2IGi!lW;3Ar6<vA}c-='
    'fl@f6b5?pjmb{>N@e3Sa@(Sja?hX0q4e^#S>cwyHF(b%~3|xm((}y?NYgJUE^SbF$>xaJ*roy_nfS5DcUl;D~C)wXjpcnP=(Cs(<GuDf9)m(CbD+y5S'
    '6Hj2V3I&boe7A#~?qlqXuJo9W(n4LH?1ZaI@7vT?iqThW>*08&b9g0Hd82BX7^0DiHH`VKHU;sfdITw60+I-!K8)7QZ;%>=JlA}!2`?(m%!{7pvo7a}'
    '@Ii4>Gx+{yLVWqX_({|gy`fC^1ZV$>GGs+-Ie9^-CNgG50y8@*RffpDe9Wkma6mL|rl!*?OgVAcAAAwK$|M>*>)k-fyps%asiM72j<(Pa{?ULMd?BZw'
    '1y-&IFqQ*DF8?Geo0%kkl~!XmvEOw|I3ADL{s`M`)2;gb+2%!auT{#6%#|ocuO%Y|40=OxZR7I}=QWJhj9D{YY&>;N>&F|PC3bk1eqCkqM&}t!;m!D>'
    '#*${9J=w7c<}XVCU4aCFA;%#~8#w^ePSAYvJfg`9XdEx*7&Fs^*+mx>==(vVFyjmC9xmN_)y|o-CE^bKe!Tzk!e3{wH&VU$FfxUQ+F?5qma>yHAOz~{'
    'A(MKNv?;A)=n4b39pV-y=ZCX{^+O`7`YAkR@+>qmh65;AWPg4U69rN(h;qwDEW{lLbz-8fk4jcvUf52mo{)Ltz8?Ph-?kY@$R{CB!eq^_Q>gH1x$>&>'
    '8B}_$`mYi;pFu0Ch-RNcBjesv<8ncLh@|qTkirg~9~nr|nDT0MlpZ}h?lU++Xb0KX&0Tv4_wl}7pkbiKIzH%#N!*4Gt%&ta{4Ij9)?+xnehCYM;=07%'
    '1>3y<y;|BBb6>GpWLUT?&`IcM?vwyrCFP_oeAXq*ruoirH?YB(?(@Pq?&)FL?lcmNOdc==29qz~`RhDk<0Tv;U4(KyvT#PxfesO-4A(=w;lSesRkPsF'
    '%WAZcXA4ZHUiPaopvi5}i7Iz4&jDcp2g(CR+@>Ql%>*_(6(-3j7FYub<Ag(mpfgDBTc1o55^GN~_hl9-Bn#vytyL8geXPwrY5Q>3Ejr3#M3XdEi>%6K'
    'gW;{9B{&;bh!ciMSHTi(s~MOi1|Dwct~q)O-?B#PvE@-3@0PIITX`?SMK(sa4-<!|i6Rud@IL|SKt*`4nh7Ycubbm=?~@0cJB%;ihRh&|(Q8=i&=*cr'
    '?*78O0OD7BN?6)b4NG!iKfaUsb)Hm4q3tep?s<*k+jP-xS2b6`yE|@tB?Wg3C!y0?nULAX2Y8k^mT)+bpRn@N=zmVhI8GNUGHzj;@k_C0L{15l!B^<S'
    'Ti8p(0zUOXc?|o0*<k_D_nd|j<90dHQX1)byipW%IWFO}KZ0-xtR8g9vOH;h5*X}tL*dWA{Pbh?Z`uCyZ=ZhrJU!g83P{0$WFU3`$g1OFotyRnhPmzi'
    '*)Ec^?F^<a>;YGkD+(;o1cCA4aL^hS(OL{BfU}l=y8(^F7bQRhs<A43I9mb27~Nmpwr}{hhy8Hi@MNc#Q+ZXXDr3UcDlsM)V*^ddIAevTd*AfU-4U4q'
    '&h#zj-KLTv76aR10g;&6U6}mmFIh$U^8{b|P_?QsMqkFqh=+a-o8qV$JC8*bN{yzv5J4_obZ3GM2V&Xo>lY3ED@MDVVoXtZh118~E9AvjJU#laA<>nR'
    'oXgywW#Q^oF4i3;Gr4L)tbMo=EjzC2W_?jgqlGlOG*Rj+IC{ZBG`QkaGS<oZfOgyfkJ-Q8cdIQiDC5G@Lbiu6A|!k8&EV0FAg7^HfGik}%77mWDm`q`'
    'HD%)5?udAl#ErQJKL3#Okvv?xS@eH8I@Oz21R(zQ@vF;nMmxwDDKA5^rjaR9EWgy?X{P$WZ_vOZ^As@DpM`mrAf1Xl<QNxRBh3S1<=Oghgby57c4}Vl'
    'd#aZT+vD3!w{8ICTI6Y!G?Zveo+J|51b2r5NrM6~Ue9Z+yM{7!T?d0DW~!G}baDt52|a+Tgj<7J?sJs(bNQrPD+TjNDX>&p?^4BbU(exr8Ey(v`RTV#'
    'K6(mdoG|OZuYf%c+e}H}l~N9Xx(wyA)lekKLS4gA&N!~8C0ys&;%U=Q!|oXuV{nF`V4|Ur>a*lNWK0ygQ0GWa{2q-HQ&K>bxDok{`9TW&^cxqs6mh*z'
    'wSZNg86nyIpGEaCWQC+6Fhcr+kppPDUC@sYeDy#n(W#hzg=M{c72a1#X)BPYZ7EfVe^gv+Yy=2gV80{P+$FRIEi1YKs2GLv+fL?oG>Zfi(%+p+;bby6'
    'QFx$)3vrfIS)~N7>;l+vWmR$Zh0N1P{6cmTtc4r~x&6og$Y~lmN{(hTBUo=jEh)=Z$J|XggWvxWIfYX;!rAYJQ&YH5y8<)Lz8miBV+$SOs_%v?rf`8b'
    'y9~+igDam{ps))np@L^wQ>l~hdr9?@m*g_ohG46G$Ct2K`^B5B42i|oR_Z_Sn(1~6R$~M-{{cWf1+;PhR{xQSm>gmd<W<}w|B(?EFIm9S86){M)qmhM'
    'Pi;oHY^*3GMkxEsCU?c8N3U^Jh*4<kDYeR3srlxj#Nuz+o3~hz!Hr%4^Dfbz5sd#<Uxp>d`#O}nNE{@T(cJrJ@&^S9gZ~EiDp2=W((~0&1(s^L@Tw&h'
    '0z6_i-?_EN1xYhxKrgB(dQn|K5A~r(CE*!_@ia|~#$+#l_;Vx9*eQv#WUd?9)p2{MlZqDSi&=)fUKqG+j3@&)6jn_4`ku-gw3NI`QaDQ0^j7?Sx1xJW'
    'pn)Y7PQSx)XZ^5PdsB!K(j7|*=k|A}EC0Bwn@r?cvpJ#A`pxV6T9iiW%10}S%CO0xuetIJMTR(IJ;+DGAOu@6kgmS2gCn6Z8v6j0A^iic=-gp(D*?>E'
    'W*@w-p|zu#bVHQc=HMl+0(IyLB>GU)6qG&n1RADB<m^Mh)*he?j?KC-HW~U8u=yxLafWy}sD$#*|0gr~ufMu?;KFWe*lF4BCR(7W1W%!bj~Xa)4nP|L'
    'yQF)@*|zs&6o*NCAh|08au<dZN*E{!Vyx08a5JtEjj*m#Lp`k*BW}uHe<xqW-TF`0QeTVLsMn;1Ojz8Y!d8WFkPW#K&^Uyi<C7@AWmouQm_1<&%wIN@'
    'kLd!IQl+G~7&nM#4LL>2hh%xgI)@nZNJY|6oB8e2k6|<*uI0HEWVtt%qk{unyS~1K1HdNBhG{c=dP*Mh;&6Ra2uJCWPS&aNC0SLlM1%<?Yq;_Nklv+h'
    'W?_8}?E`9P%XtMzW^WGCTtDtpTONAA(E8&Ux2yO<==%u2CwVFJU259b8{j0b!}~(`UL8Gm4@x^8Gcgpqmyw~E&6z!_d_U~k`t2j9^Db7>YUShewTtqj'
    '$5u6{+`Y>bI79XxN9T1Q3c{sWu^vl?j<&3H)rn}!x+OvG=9_cicmG|{<TZCNc_6}9*K&#xbY{Utqh>G+7GR4IIx+FWcEC!aK6i-u2R7Xxbz*a*FkWDR'
    'dG_;{n~WZw@CluJuZ9uKs{1u)4|Coh3T}vwACZ0g?It^X`1Qj-|L5obEmH<>uT1iPRUyp^x1F|u8SzhkiQ|NED^M|=%yk&DSa}OraG`A|5un(($K=RT'
    'Z;yBkv*2KLv3>cp?7a562yoeI4&j+X^Ep+T7cODXDl4&4g$Tu4E5XH2&v4Ul>U{cLA?imt|J0N^Er&P<U`zzCa{$bT&u4|!z7T=uDm^CO)Z2~{fL1lk'
    'DhCjci^#G>w4a#4MXa=>iadqnu>T$2J3S>-ke2EL*-=v;SN+w}$5o<Lyo|dhq;mbR==&dFua()4pT7PCoA9eISoAeokaFyOP%`tRL@807BPCzVW(@l5'
    'Jyq@8V+vY+(ju8Z^5c0opqoGWa#)vYbS33U<*z27@vH!-dO0Ysbc%@<4I?e~Y7j9#lw;35rdwdOgk^wJ)W4=WTB1CaH}XXJ*<&avLyV&jM}8VgT9~(!'
    'l4nTD)fQ(Y`#h2o=ug%87)fBat1>~#XON<01&cwx4vwa5ci#luDXgvX1Xeu(%cI2c(!?2<xarpdN&+Tlz~aQ4bm9w1Y~B_s2GW*gH3P8e(M=13XkJqd'
    'ZJc@~k#(&b;rdrOIDv2>I;^b3OW<IhG`h4orW7$*6i}4DT`7#(sX}mlH8)`>00(dTKJ%W-D=I$@qw@RK*aBo1?G#Qnw9W#C0I;STn&Ucp^1F3&{ivKs'
    'xP-gxcB4YI3J%M1P}u%!c+BAMAn$|t<&QZN=|H54IU=12-ry;biEpxVL<X}c)#Ob!zG;@o!g@Xkd1D@2$U2i1@m;e6EHF)2sPO>HNL0=dl}Mz@ldECJ'
    'L{%iJ=7_3H0);v0V#<AGOmj`&*X^m{o|1lWNu;UQUND|6PKOUO&1b^(LhVVAN^N}0rmrcQhU1>{>2+oEa9!E>Re{s;r4=u|s5YHld?Lo06!H>M>M)D3'
    'cM-~1U4r@}!hFem0(I!o0@u|04ufC#FT;+@0^>8A9lEMvqdaIaI!(%$;KbKD+0E`KB}LqV$b!d6=<%UcI#guToCJ0Oh8c)U21g+4B|^$q)-iQCUL!*X'
    'G+R=#YP=%pM4}G<AY)oK!3?+bw*X?q^B^ZP0oe@+f-a{B>kfX?QO$PbmqX4sUMyR2EDvw)xKm+C*k>QC38O^`^o4H&X~riw(xcFQiPpr8VrJg$<cu?F'
    'I+jKx)HXk-rVPzJI_xeFKAAqN&LO<`IzIw!^SB8To*c-Xk+1&K@zMyCt4_gFP@O}NIpV_TotK|!XQrMdstaSUoJL`M9rIaYUvTWw2z1;$R?dJpefZKC'
    'xY!ksAx{6H1q~YmgLcQ=uEXjJfx|9#(mPU6rl>@C={`!BSR~&)z0pPRU7;=F!qb_q$C`-=b}=l8Rn3uMp@l15znUmizIl=`RS+*(n+-Zm%uWf7U;Rc4'
    'Dw4+wH84P0spLE!9dkNoQ*v4O)Gio~QcsZk@H)3;&&T2ImN<^5fZUlq2T(c!m?r>yboG?p#|R?5eS)4&%-KYQBn+Gg{RH`eY9qOW>$URpdG%#e5Kg`j'
    '&e<m6s+;+|GA>k50p<W>aQIl6wpN7|s%$!>k>=&?4WVu5DV!BmX}`$27!LxwC+>!R*OA3gpi9eLYxqBo34m{=P?S}989<Vu1$2b$Vf6YCwWLQ&JfSux'
    'HqwkpReW)EnY4UR=i@fQPd-QpA{q96LVi`EZc4=Ch*(t#9}z3KW9x#9f-$Dq#ja7)98u_C;KLTOV&?e_yTdk=ZoF^m1)feSzH&$Ps^1gDI4H5OHl%oQ'
    '4l_dX`qQ4n|9x!k>Z_)HQ-g|z9PHINb{1~o(6lA5{J1b0G3nU?8Us#&4Lva^^Bmj7D3?o1O+*}PCh{eRp<)e}^}54am=D18ebrCwq-bqneSv<Ag>pqG'
    'm^{=uM$QB>(6`?{-DDarklgvpuo@2)rj?=YeQ8-^Wr)-oysC$uW6CMI3!6#clK8kpMU@dkyCyEt1PD~1b>!|!vpA+UXIHl$3nqu0GvhcgVNJ3Fcl~CE'
    '%?ljIiR4A3*3VI+o}0W_MpE+}sb?#Xvk{(BR3YSBT4yR6*3pmo=-j2(nX=j6KK=CXAHKp72VG(sS}%RSL<-ke3DD4OFv^VzZ@bHgEWhUns)z`39)TlB'
    '=E+6!B|4JF+<<34!gxlBjiQ`=o(rH;VQmm~$_H1ou|Ft|4J6MnH-q&NIHADO^7yvPbe=a25&h*?D|9HX=&c)YlQa}Wya0iWQeigrj=u7&f^M#@tBpX_'
    '^qu=G0B}fR>$+^A?}d@B7NF7fnRS7QH?QE^tsK`{f(t9uZ*cL6{kt`1%96WzEJwWcMV5qVc8h7*CZAcUvFYp?6dAh`fMJN-rlihA-uHkVj0#U5@|f1$'
    'w#BTH)+g`_=br}FBMSI%j3VpTDJXiAMW|V#of>`)dVAbsITRLEq;jxYs@rWlM&+7g&X&xfqK+-4sALGmpzXo=AFz;=D(1VlEkF}x6`h$qU-r&BMd>sB'
    '2G*dX;D2<s1~H}qwY^U;M8r_fV5D4u1J}!<2v;r##ymlu6t;2HWe_!-JMal=CmoMoUdES~&ye>?&^bPDBdL9cbks>3393k2JwrRH?En%e{GTX7I{;_L'
    'zCh1HG#{dRn1TfRg)N>~`&eV~0!$HMo>+c5I%5d3!ICN77n+e=OK35WYnc>Bs&j^!T)2MdA6sV+aCa+T1N06Q5*vz|an9&*NlEU*784N5*9)6xTtPRh'
    'Z?h{ZVjElog~7s|RcPUWLI+_Ci{l$=OdHFNt&!1D7;knsE_~@!X(w1dQpk7uiC8hRW`dDe{xGKigQI*t2MFfX2xN!V%SQfFcyUz)&mu=mgvVnrOmipo'
    'c=13?^<@V!9-ZjWt=SB1y$P#@=^|T9{&@s}P1(^c!7$R;2vb~w2`NM)ln^1FOYZ~D1u+s7Sl=h@0I*;9HT9Cmvi8}gzwb6(>*^O#Y`5{{_?xDdKG?R9'
    '9OT&4P0d_vg#icK?qSrD4Y@4V>$V{1h`?n@51*kl7E8r<PQ-HAB&fLf3P1(7k7rWm@SH?K0sE^j6Tf?<LIVaxx&vpcz%{m@9uA;+`2rNGDpX!l{ry;t'
    'fz(<@!|_Ucu7I$8=X~Au`fV*Z6odq8y4g{N=ftd=tfmf!Eym_JVc;fC3>MIGIti1L2y&6UBy!B6Ds`;Y42mZe+KlzGBM1qPdtYJxh+X3MsIWMvYfG&&'
    '*;~+r0$~{btb{B>1OiN?43I1V)ZQ;8;7WYJe_(+ypt*K6$9<8{AYJJIfvFddM;Erl!;6O9X(vzL0wKm`NE?va^*{fw1~c6-LLYYQeZM}`Z`dzb{?VLD'
    'Yy^Q>L4VA?QbJJnZ`rS(|Mk<apMLu7_s=((szI>&jnjmhTUYWlDLkJ07o|1qq<Mh<I4u+Z`nEezHF@n_5JUq17R|dtp9-uC3rud|+CA}Wym(|;^I#fU'
    'mDt@RF58u-dgXdUp6EYub_2%dc#HY3&1;@n7Y5-8LW?<fqH9dz^R><~(<{QbHw&4$Zj<f)b<_86I|};{X!sNIWU~;AK@5j^R{S{zOByKi@Bl=T8*$I#'
    'b+c@=4#f&lJTA^a{HHC4ZkJ&hSG0FvkyV+d0yR5qWEaFBPHmY0#pB_6nk?#7w#kG-9h&}8G@D7`+k(Q)jiLX#gqONvn$2(z8v|Dm-Fbovcmq>~cY5Kc'
    'S?;u}6+niLGS8m?e?%wB0yR3Wh6l1LoJlYK@y?ZT6U#Fws@U3e?W%s39ULh7P>Ikegn!Q9p||@0qTn93FO_-O)D7$%>?Aq#+x=_UKRvAy>_fMUTR8Ho'
    '=$RBAm<|B`kps7!mxsqF)k#(-dt-g$o8kojNl()lhR>1PcX2Wg%?8#4#^cD{o;m0|w6w2G<a%>QQfju-V5-YHIBy7|h!E;Y2!6YqfhZ${ehR{`*BHY6'
    '+hkUr{FE>qp6*r3sS++@&rx0nFKvM1mGUerkF}_^F+n$SKLL$Qfuv-4bC5*)6DQM6O3D=>DyB46B=kMz?yXUvS@r9F51JsBU?wS(3R@hqQ~52b;YwSO'
    '?tIvzRqSRt1ACo<IU~NSTdC!>p5lqnowY2>=kT1?=EA=SP)(}Mg@O~9Y??m@I9vY@zB-IIDcIx&?90huJid4ipVlKCZZWC3!9|}aCblS^{r*IE8O^Ft'
    '(^0A-j`G=AaK^H9oOU8iNUt3iTDwcixq7g3_7>UM8(>;{Y8HBBO1z>Z^KFscEpsK$8XGp7h4Qx+C=eods3c;OUgB8L9^J-{>+c4WpKRi0W@f>4ffKJV'
    'wr=eQil=4*->A!yEm*m8PJU*cLJU2;l%32I8k@K$b9+u>@`v~yvXg*xT!hl>P6Q(s9on$0YRu1L=^~UuZ(L_wuE~QK1ghcwRMb$r2%UJ`4ZWpSL61jl'
    'uL2jlyPCaf1}XtIQ3@foqB6M*blxeWvb$q@&q@8qolw_E5z7wbC5Qtmgn<ZYS;`>8+k#ZeKONzC9x&H|(+a34))9LHB4rT1LvFC-cy+2%lUHU8$M)C6'
    'wa8%Of-p+@9u4Rt;BTTxfd=E#o9H7*VklxRiOWaSu8;tujq#0_hHU9M`8<A-p|U~S)D|De5^Uov^2ykJSzBwlb=@Uw@^+D>(<goG(?M6RC3l^Rb`>gT'
    '7c+>lWN)loa6HtZ)w?V(<@%iN_IOx;!bsfwk<!gT658G#+f_1>1HnW#1S?t??-nXbX=Z@aa#t>G7~wFHput-CR~WJVD%H*lrDORXEu0KL0-dez%pOB@'
    '53KKfDR@_gA)^pRo=je&;})f>;krSomD|F)cg$iV+UsEB9~Gh~1<1bMRBAXa4D;r*dP%lvinZmc5SZg(QC-Ea=gAoN$+IeENESr^TGzyLxyFzS;8R?^'
    '|3({E#vl8Qnqs2ry9TJF>dVpt!gsan+t}(_P+CnW>rr4ywdC0?hUO&7=$i#ud&`*p;oc@gKo-2MN)U$jam*Ji8y>iRl?+qndmy6Iu%Y$d(ecP{N^Dv%'
    'p7y*9H=xPOFkl{S<SM`?7idU+-5gu)$H}`eW(-?yWv;whE-Fs*CphxWKRqJfZx)RfdFVImN5oa8;8;K6rFUrGphYQ8rmG<y-LsrY1N%Yy?smDmSIG1+'
    '^h>ygv2k^!qCu&-(^e@Jk`mmZp&~dWRZXHc#?B!fEkdwv`ka*RYiS>3^&mCX6Zto9I>0W9*<U{Y{`og<rTFQ`52GAvOmYbIj+GOq0{Gi*xlh(g=%`qd'
    'dl%%{;D*3yjntqulid%*?c3Y7-TOo!*p--ffrLRtW#dg*M2nI$NH`=k{&rNvdIxuf(LaBIk45#Miaor!C!a6nDf6nTQi$92;}(nN+}^UY3k2@AA2>|>'
    '1o!yQ8CFJTiwfor1Xo#6UVu3qF?#VKIQ{)7&~a}=9-$gd2}rrCBV1GOd(2s5{~r3$Z{hEYBi8mfDf#VYil==}g38@lJjXj=NI`>b<)#Q)Rl%xWYVGdO'
    '83g>GYc$Nlca2c6YHUSSm%@2gquIs&w0LaTHV<b8xBi!}Q+g>PeK|*uE$&UIV1izJMQ0%Rtc6kIb=|PP6TJ$Hc39NM?6F@RhlBIQ-;8!E=70H5z&aHd'
    'fGDm2t^B7lXMaA%=7c&kXfe|j?=;%1<`U2bZ3?TNO{hOlh9W|vISe|ugwX9V*(#1hB;CStV{+>~ZW5FZaRFslKkk969@Ac7_$j63@-UuF0|U4;5R4pg'
    'Nqy@(EUwpK!jsAK^N<al$1njU>;*u#Od)y$6a?}Rr#ly&A4#%uKHB@Pj$hgUCOhdef{4Q6cm&}D{auvLp`73fM39W5T7;bBkrpC&-*sgjH%1)<Zmw=Y'
    'jUBVP(N1{@vO`gZ)&d;U<fyQU4BmnLn$8ZFv0aX(e&PDgBR?qp2avW^Kh`69U&F`ueHXzEa^R}gAq~*|^gupvpt(CXZyysYniQe^g!HR9qFC1JP3@9b'
    'H%ky4icmQ$*K^N@d~YIo-wz-M`?kiO@sB8g_0v;9e`#)&3Ti<m&D@9Ff-bjwbSkToGCrQ0naAItD9wJou3dUnH?)>0a_*clC0loWI=i9rn)^9@eAA~+'
    'XXR6e;?F5Tf8RQS%?`Vc#E+n*`0(YY-#=u#Up{>O{nJPI@$JjcAF{<SA3pwd^pYyRq?&oj9*YY4<N+nB3R5pbeQzwSlCH`TuFY)dR$WX&c2#pGRi4Vw'
    'rU8`w<l`LesI|A>@(FbjJA#YioEkLAH`ART4o>ik3ExXiTW*avM!^(2_<Ya<I2|$DzVpHT5u%8$HC3KL>`-3abUqi|ffN(tjn=Ulwlm%g7|rs-5|qiT'
    '-7bG%o2bumif{KB+u9mupCLk5G}KRGDy4zcwd&VCLJ*%nuCf1zuM!~qV!*#=&pc{>)FGr0*TTkQMvtIK8c6+JHzEz8#v?ik9rJXmn4u#bklv{St$;bR'
    'CCY@NwB9f$IH5PvC%O6|6ou2A&?t8X<YiO@l9zydaKjuzl_v=L5`yn7k_VG5NU0Jlff--?K~_#k-otHiY@60uK!l6Y86T*OXeWW8s`R0R!kcXNu>Yy#'
    '=g^P$oz|i3toDvUe7FH!V`!Uxv%>=Slf)em3}kqNts+ss-5-Wyb4b#2cu1}aJ`p$*G7@m1&mk#u%r7Xb%7o`t`~<XPVmY}!ZK+cN(=woCR3`RCtUNOw'
    'Q^EYfWsvp<!m^Bo@JG*h){kVazTBUV|NOs$V=;b_DD(J5IuLbROrZh!XZpO4e4<5(&!ZWzy|35P`4S0w5ubL26hl9t=}VGnt85_KV$bo$nF}mdDZ|`b'
    'B|}9vPQ>Deci4=v4mp9oWaE@Lq(N)XEfQij3{HrX`SuEB?uv;xV(PRvDiXQssUyQ^oq;`7p(QmR9Ht-o4KDHyEk<{2okYHmrVd~*)vaSau;)W#Lnm0b'
    'LknVAGxQCWD#UFWN+vpU4Mj<#;Gdz|khp$?EraMtKc7V7O9&O+6fkp$n<i(=1$|p361b<wbQX|_azEUYz8pX3TP%nI82SUS&5oP1H+a2cv2zeTKq)Gs'
    'cIErZ=Uj@+^_Ax_?Lj7fd)Mubna<al)+eP;+DzyQR;HM)n$TIGv+{(k?DeFNeNn_d_H4)^x^??-bd@&H9=Oho!W47gj2h+(Jrvd?)a=zIYO*e(DKG?k'
    '-q+nd%!Cy10o6&C@E_jg2Dc?hCoH@>D}El3aEo(iBVf+ftjDH*7k5BJ?6f5!HvGe|;Z-{IT(AYlirj|?E8)Hws{+DmZ|hL<ro^c60<C4<OHh&6%FnVZ'
    'ecRk6@_Z$|%hmscKd)Qny=bkYQs=H%Bi3>O?LmteGYIEgF%enKfGCu~ZjPIpTY*RvL_(1t%8$lGD=d%Cp$ee(zh~okixwLA`moDD*@GDiX7<v#OTt!n'
    'CA6C!w0A0p>Gr;$XsirrM|yK!m7xbkc>4qlMIh8Mu>d@5!u&t+MTEgur~vz+35S`3ylZ^!jpeq$XPE3cMhhKxXrZ)$zy6#dZ77r@3$jysWfS<BoPe#V'
    'RMEP}MBgD9Xy_()QUH48<E39^aIJ50HGP9E;~Vr#->^q1mE+c?Z^4&{(M!<48Opy{GF4oB$<Xe4EM?V?ULpSaGXf8}sK>~8Vlp&lc7os^VMegH*4xz-'
    'A_N%<F7rDNk&I^ygGEmSY@itPV43qxVf;p7rv8wb$cGhMwk&4<(usAo;gj2df1#PfW%W}yd>NzRrBbh(+9XZjJz&9r<<o#3X#&LBsPKOB3WV58K>oux'
    '{E-hv@?}02!?li>Lr5n1{QJcG_i?pzlNaiF(D;p0Fc2ik{wBSCTa{`882Ad_E87nE!R)s3NIqQ*k>qnoe(KJFq_<^brVOhtPXV2MYzB<^T9F-kl!6gr'
    '{)>_;4n$bWw`h(FOsT@WjJQMasAKdJ0>8>kS=vcnhxFx`eRJF=P+c2=JtmgDKZ0~WsyQ7oLnF1)r3KGo29F(M;NaABaNd27exFM?A)0u^*yv-;1W$M)'
    'G<-IPWy}H29QU}mRk};Y{A<&%m-QivaPynPhT8geasUFo^Om^XVDjsV3fxYBbf~cEVqMRH&i2+MJ+IO}*ido1ZI3}sozVImq|pZZa$+&In~Vt4h+u)`'
    'o{A7EttUuqVG1HHwI-7qi^Ny4)iL3yDwcHvoN5;5thyXRsRW8X4kR7l6kAwM|2&7<9Jfn6gXm|R8I#*Jxd;u{IO~{!xYisdIxUBO7p^ScHT5QKRl_hL'
    'LnS8$@g&MPb;`x-*x2XWnq>1u40l4==a8OD3EGv>G1yD7P3G>n!yLX%*KQVcg+Fw@{`+QUWEq<WapXKbPaVmrq}s;S9mj#6vT?mZq)wb1`Z;pAf{<q7'
    'pRg)~V|(GuVMI>^n!FE`Ub{TfMjvC4tI&6hb!^08qz)u(2*2#tYidB^6;>>-bezKZHo>rA^Ws^!4H{17p``I`w={0qXQ4(A&IRj=Hc*6=+g2WeW;Sc@'
    '`!z5~%+&^VOD&JiF!&B-4ErPoc`lvf5}_tYc@>eQK1bRoP{3-3$5{{#9bgdkG>B7r&rz&Plw}n2G>T6MK7)~p8wfcVI}7u07}`xsn=b|6H8w?2hT5-Y'
    'wEJe=Y8XrMFa;?q5#8hL@R$L?muK5!dtX!92WDVkqzplQ{9Eo*Qb3Q*OM4N+j9*>2!X7||et|LBp}!_RN^c%jk};IHj>WX3>uLtC6>794b{MyT+l#XQ'
    'N93`QP!ak&lrB2|U%vhAr{90e7N36k<<nn(`Yqf1^z#Qe_1a%=*gsThV=?F>I5!J-MIOl_>9yi!u_|Ptqkc+fkeu_2?zVoP^xfI}VqKxae;Q=3M5#mk'
    'v(lWkL^z>Zf+({M77Bnfyg3Z}8yB4tD))O<3v!=~)IIs6VI;owK_=ve>kasFxlMzK%3G>5hOshLByaZ(=K-O~3LgS4t$$5=Vp<ux2lG#+8@dJN$qN`N'
    'N1ahCUQn3ys}XS?iSt=v7ePwo3Z`mIE+e^|C3jV(9eFu<yN={~mi(;xHUVJ>&KO8VAT<YcNDEz!1_fC82usB&QTg*&>w53{?$w$5Or`92bS93@JdZeO'
    'pN!tZF`SKX_5vK%qptTLDlO}@BpvC2F*3fgl?u&s`R52MmNTehU2=Ry5-f(%li2CB^nnb5tK1;ak;0H0kc!&jmgn3@=n9N^{P4@~nS}pd_S8FNUDLVF'
    't^n=T1i$o+hY;UBk1svGqCBXxH<aq^4I%gl-cY{a4f>fkn8_PdRV8l_$s5wd9k`3y?Vj>%r;q^PGDbjTG$9twa-bgm%WSCERNNh_?KX+J9IRfc7D4W1'
    'BYjY_|M?`)DKWy3SHtGHPm&YP(c7=NvlY4(Cw}5pO>t>LpD<g_e4O+VG^Gg6G0EzT2U`qBms^M>Oc!;Ars;D4TLj;iWMy~^!0DTz3RQ8^@*Km&L2#I%'
    '$B1!$1fy&ux0MEs#|W9bI}2|GrvRno?Cm|-=<uCRfv&=Q@-Fe@F+SWmaw?xXjc#{^@2g53dNXNtftos})_yJU3tZMMP+#ZDN1TN4mCCi)#_*MjFGdea'
    '5zLA2I&*<5KoSF2h*jyxwK#;$L#&S*j40_T`WQZH-*1O5^rHtkXT<z;5n+mX7+1oSN^pWwS-l!ZCU~8QA``|Ko8v$KD;nS7YcTr@1KUp$r^m<PKl+qq'
    'o-cs%Uy1P4)d19R6;_}aeW{D~aVP1~Z6a}FFx1r{zsA>KdHx2GQgp6wA>-J`=fi<}Uw(-eqo+oW9skbYb3QW;f3$r#4LKR)Y}R&}poT3Cf?z12EI0@P'
    'bfJ3S>NzOn`VPt^4*BZqPrsv-gGGIQ{P3eMazPz1utgY7L)ZjHkKml!$Jbz{<fI27VKbmGYeN}fQ_&Y|HjQ)5Sie{kD3aesfJw|fMm$neA~P)jWmMpc'
    'UZ8W(9fiFA@W1>3SIFBm9XtjM`=kj)UF{Rvew7{S)g++J#%lPw5b<gE1$Xtk8k1R^Wx}VGfR=I+6C@oQsOdsJfi&u}JqEKN@r$x3l@*~td6}PK?b~H2'
    '4TrNjCN@R};-Z=-PAiE|UagDxYB5h8>UadI*xgANrJkokUo7_0?m#WermeJvZ~98DjLV4N<_&IB3fZC@7fi-oG@(P8#+()`k!4~jQYB^}(#;2Z@9`a$'
    ';4fS!9ZXWf48yz(MI<W4GUo(_cJoZVR@%F=sC>+Y<l4QC{LnV&IP$F4$7|6HfChyPh0EBf(%lr8>D)cQowYxZ=l_K9nOU=0udbK3*SSl1X@)|EPAIvx'
    '))N!=K~LC@QnYT@ZMWbum>|<`UisSk1-89a7>{B@Z)&z5Mic;9feZv<++7lQge?xeT{VD|WdsqIK^!ligRr3nhUJ>P7a*AP%_8^`|A?S3BXDV}KtCL?'
    '(P$Z_vN$6h&t#a0tKmSU?AYA!`IpZh{{7o8zeb$u85~kGH)DujKYaO1Kq!l+5uEMwY7sG-O(A^tE|my<8Nx0i9NdN~;@)~@7`(L>;}7}r`Rj*|aUA7y'
    '5GcVk$#>kt_V$7{h&9iwi7S1-NdZJz>Ih(81n|x&4EJl;5>=#rW+}aVe6>ynurPTjlq9WDXAqPa=1v4$YNS6nM2}G9p^b?w^LYe6hy|=<;}TsPQ{*~Q'
    'il-?-KX4hcZ<j{K-lr;;^Msh00^0JCjJ5m$Z?OjZ^0+?)9GOQP^*oNd+h%Zr$41T^h%+3wrO$DEs|`#BtDeW<j6JOPHNk2Q-eXhbCV#<*4*d>+uwNz|'
    'd5=*>jPiRhc)lHqNJgT+2ZaJ;8AFPQVZH}rT2u_ST$~v_bK;)B3gQKfCw`O2zYinis>d-Y>*_9spD)IX<;R-$?Y58hMh+6CPBj_*B`s>{j&9%7@p&1Z'
    '$A5qR^ld%t2LCuWtmDUB5`8Rc%1wDdx!82ajmx#mL$V!PLsdfV^l#1;<RlRHPEgn$eep5p5Aj6smj!kYMk95%z3(u7F`<%yiVc2aDtUoQMJjPlr7|^3'
    'rNd`A)pJCpFHo8IE`3gA&>}rUWh0e3r?N9t_=J#QJ|qYCIy#rhj2aL;gW0wnHK<5PeTEs5>I<Z*oFfgHT}UeDNm0vy=Vr;PGJdRig1nj~SCx%#SC`4*'
    'xe58=`2M{NJ3UW!T~Zbn9~u>_TMqTA--<|A%+m2=*`i*rN95EPH&P2ZOC8$hp3}w=3-vv;aH+$gP2QZ(Q}c7KAwdVXctRp(N%%1z%l@w0-zBqSW=Zg&'
    '(eUtJBx_+`M&-6&=#|D{m_*k;pw(l~>o4a|iwYh2sxrPY>=;a?1X*a``#4diGI2g4k*g<y3VJeNVKQHQI@z#!1Sp>YfDw!V?m(h=41B1FdyoEb&ds|m'
    'nxrH9FXDc4f7(#(r87cmZW3>exbocf<ox;L&~G9^GN1gHa|)WwJXwPHF+odZwv$iSf!e6V!Nons{Z-ChwQF~~WDerk=npf`ue%NYe<X>oz<<yGI427S'
    'A<#vp68?BjH()uW_&M=d(;v>C-?!_wA?I?Si5n^WahC2rnYJROxu2N^GB(5Bi=PvDhX30c)qYzyZ{rNugoOSwZdJakcQ_+Hiun<%apEs`zR_mc#7z}v'
    'hW;=!;$62)rm~0^8UJNd&rC~GZ<MOMenf2_m~7V{$ob*Q)qu9$V(Gvwr|+c=d@z*XC}Tr7pB08dC{^wR7iWxRBZD0VJj&3cr*sc<U!KrD>UL<86d_NS'
    '^!9PKt|{H7hV{Jf_RXY=ofKNfdf84Bni4-w+v=TzL=&gY^|K(MZp*j?j4P2!kzIb%I~ek)#xkvJosjv@&YT<}loAm_pF^<2FBY<%2l4JWDZIgQ7Q`j)'
    '%wzaM(^7b|<t&EtgwI0QU^jK~67vw8y2lV1lTCnpC2jnk@-mR`(HF)-ju6<(@Cn)kW3torJp`7fx@#X$y78am<2m3NIJ-WjSmiELivD%$r*Aj(-8S1Q'
    '3|R~5In8pn4e7qlH9kwG(_}?sBo{37q}By>I?}siyQj9cR8;^I&NpzLwmw*zYFuKd{FM&P6V|t@T_{;w8Y5D?HC~lywPrk5W-!8H8w~HDjpGx_cQL^6'
    '>C1RDwX=hKyTU;lKS+}B9%4N%u3FUZ>O}VAdNB`~v(RoEBW)FCSOkhKXz-6%S{O?!Zn4@x8yKDsM{0K4NR9nc^i$Mq6pggn*ifC*x{Gv2<=EOocgQe-'
    'VL&g<>bU2Ssqp?H|M-#IG?PxDB4qq&m(%IWTe6eH>9H-`a3O!5n}Yb~NUVs3X1tk9Ds(1T>~QQ0?!&8`n+hp0Y>It#zS#T8X85%4)4;5m+{aEh9F|-y'
    'iL#W~qc^k<xK%)aAOcTnBWZC?ddd_Nu)APUZ<+;##;FjehkeX(68Z+U^GUN3p8%@!kPekA=i+9%kdrPeU!v8uIKeWCz7RHMpQUfxA3-$p4oX>m{B_zL'
    '`y{fr-pOGIuPlMb_y$6@67C9Nyu;AJ?d$=^AJJ?cT-Z?+s=pB#uFKHg)_gA~Zk9A6&<IT~79Gk;YiBbUdTz=WT>OuB?+?NUu-gilJ8a_8<iUe03U^@<'
    'DOY|NVjo&pj|$T)AFqWG)o24_K|Nno1vtG(fZWE%Y~;@RR(H~32(_-b$22$XED|d3;B1EDdp`h}!xhD^8e9^g{rIl1XR0h-O`(m7(CQreMW~@+rY^Dm'
    '2h4V8f>=nW=sMK(@qtgns@?b=-+sD;kW(D*j>$_`=k8@VI^1d{1$6kjhp>K?lbjf*tCv6@Qik0<$)0g^<&8XS+I21@m$*zOtKGnKP|uKc4TqeH)hF^('
    'C`SSiA<V#43}=9e2Yl)-e3NTe3<p2SqOVZP*vncL$!_*HJr&qu3mrDSUp89WafUorC-fn*rBf&gB6egSj;=rn@go4nPiwb9umg_q@z)#5xWGDYgz5N4'
    'MdiV<`h|Bc8j+6EYeiZ82H8LCE$oQ2k?J}x#<S)!mbk90Q~3zuPe-a{#yj*e)Dj3K+~^KAGSu%V@?;AQ?PGj(XnWHSoZRdq;^FXPcfZPo|I$@vq=Px2'
    '{dSK<tzgmzZu8_4VS)R}Z1MT`Pq{7@V<hG4LJ8NqVIIj#S`kUgq}8Z^`4!==DqWg%Kw1L~va@V^1l5H~J|4Hr$L!=Tti}4+hMnlVLIc@$Iqn_*?wh|{'
    'W%7pV^Wi@+=;amc2iLg1IUI*SrBE`3VnYLLYI-Qx#h(J?h#?&qp8M$`7ce<*-=S$ODUe3Nlfdxdze9Wh1tL|(2O@<)qDmGa)=D{<grP&Z`iQ)z1`1Q<'
    'B~~U8)j2(Bvq`7cc$8;88;dVu7I9J?&g7Kf5JUnq^A2iU4pQE-O+5{d*d9+NVKyRNS5MIIn<EA=;e_G5A=NdK0WSx}S(;=q;zyS<v{utaI-+H_t{>7s'
    '*inz`OP^f2`oi%Ap~E4*z&`bYpd31l>g*MhEbuZ^fiZbiCdibr&CCO|fD<*DN_O@N3p9p(o$eI;OL{_{+Rs6D)rIw!hc=^?NdtC6vzIWUDN5lbk*f0I'
    '(;14_?@r>^-<>_h*P`L8C!Gd=NU0W0b3DA{s~TTthT*CixWo_k6jIy)mG>3fLX<^<nAV*kd%*0biii~l#fi##o{g~>VjeNr2c8%sUAmNmq)Z64@{x?^'
    'h<HX_(b_r$43b0v^w+sfxHu^Mc@+{0g9Gdg)48ReoJjs`i=Umx&zAFmtR%i{CQF3$iZ)hvo`<Be7tfCnbHS`B!sf9T;X*+D6PU*^Ha@21JeWHWFWHC6'
    '=Z3+j8hZiGM_N2}M9{svYJKXkEz|FGH!EGW{M>i25*T3>N&zdGZfxL;iymdPNyuBaI)Y5G^A#RT3v-3h$3Z=3wgT7PrcSyI!QF*{?p<?7BHp??#Gpqm'
    '3ZO_d(_Flh$@xR5kZF6wDucdc0h)cVJY=?Oc6Hje$6{&M(6)!q06@W&dB7q=p>q#>$6rH5M@4sBu6(^JEL}yE*szB)GpQ$Dzo4AQhfJ5}xlTF8e=F;n'
    '%2JVWkz(vOSX2E$8fU|WC(z&s`n0J+e>s=ZjLoDr^@_@(-MJz;cN<L&IK8o$_>$zZ2izi_1NE>F$Usa#`%l@v!`f=Jnoe?%`HgY4nwM)VzA!E_tWZ-5'
    'wQAYX;}Dma$GBM>xoO@t2$H{i{lHyYeW0%M;$R+nXr;nlABW-WG}yuq(@4HK>y8X10I+;rc}@^Z(fi^ZX!pClm9D;ea|ioQ$J%|lgP^4Hvp{?QrP9PN'
    'Rgt29@u$W*U7@5$p?MnxEOd`;SJ<3nT4@QWLg!2EDykmaVM`_L4z+hwaS_1Kh}ZvYjytYvvTL`&?BX@@3R|XN=w>4OGCtMx<*rgqM=c8*D$2b(QXL#l'
    '5-PoNAA|*AqPdJ~Wi4wgBZ#^|11B8U^V@GXV=sf&Dx@+PvYjnkszbKp`Zk~^w`tjTC1ZJFR+B{@)<|&WPed4;g%8EUfBv_vx9(#9;Ylioa5f1_<sggj'
    '9DP(lXAtt3j={1-g0RHY{B5)^&(P<#)PY`4=-d4*Rem-==H$l-bu517Qa3MAk2Kcg`{UcCtfusd^h=WB-1vRX3WW-V7$d}fEPG!mZa*bAAqt+=oeSi*'
    'eSg3>hcS6k*fPFdKTqx}?2o$UCUD`v<(%vE92}`0k7<lr@@Zue!Q^EyPS)8^hf<hGZ7x%LAL-e|hJ6VWTVAFQJp(5WxXRkV==JQ>-giqf!jcWA#wNIX'
    'O3aYi-+uY<+sA+T@O5I-(!m)i*blXAU3cYSRL?cGuZXEQC8ysliBBH+x>@xbtf^2p$8|kUr4veqB2C=%2HU$0n|hiyx0RWwZp?{HjY>{YbSOi_Jzg%M'
    'XzaVh7<!6f^N1mzz_6LD90tqn%cdy{Ctz#r<w4xx(}%zO@;UqX`s0VceaQYT``1r6(_?0AUD_@(<C640Z+uKU5BSGHAF_U3Bxgn>k2`ar%M6(;s8XTM'
    'p`!a~iP=8hx(G7Skj6@i;2Pqe%Z^J`4p?}`b4VE)x1qmP_;6tVbC$*?G!{K#Ub*9+Bbq%3caBbJN4MPXv<)q`G{{@U99en|ByZFpoXFTYIgXOcV<PqT'
    'AAhHOQ_z`1A=G6fgA!U^q|!<9A$48QA*|O{1!<Z<b~{k|IJxp%8l%sBVQsE?y`lokAzSkxC(RuEH6*R#kfjN1v{Y@d*C^(0(%z*IkWpiY!=ycjb@lBL'
    'TAJz%`ZY}Hn$p-uNtn*#Q3IE5vCs`BsBIPlrbZp)sC`Fxld60-&#JP|y{SPr4DvoQg5xJm=S1hB_`2c#n@_+|n+q&4{_sex0Jqc<$4`mG#S}81j<*PN'
    'UYV!y+bvw?8*T{UL9rCE$j3Qa7H&~*IH*{tdji#0BqeiuwwZpvKe0GI`wY^tfjiZGe=<UMOvu2<f!Fb;him#rG?CpszK7f3pe8%h=%md^I*KK}5cNlD'
    'vEVj3zZ#q6Otzt@{MUe(A}->%zBK_Qjz+<ijFld}f1s+;$z5TE#pp}KKCtiLihca^AMlf$p?|Xeg%0I&3^&)9pwa2N>xOPO8q;*2hAQT;Pl>4n80?;d'
    '#Y`{Giqs<PaV}ixX_F2ve8XZ3mpTq4i?eg!SKF{o{PP?srWmtQu_g&W0ldo#ogMUKjH)W*81q?VEZbXy3hmm@gdM2Px&^_NXv2Q8>8ay-nosQ^xLReT'
    'RcEyQJS~+0DunT;y%@i@zwJ1KzCK{+*g<4h!rJ0myh;b}Ks-xAkm3mt<v`R31hn^}veIkQvS9nb6vVe<h1-M_01&t+kGI9a!MBwts*f>6Xa-bRQ9X@G'
    'vIJ>&`v{ni0ZpajAoKi5zy?cw!AHB$x$KY=WUQ;{A%g)2Am#T#N`^F(&t_kp;Z)4aaP4|ew*Y}iH3jiu@z|MOKM8a@^zZQ?5nk@_u=K;V{_xNL`T6Us'
    '=_iMhHO`E58TyW{SAVg4@ce9De1{3?I+4E^(S3f>!pv=*(1W|M^67bmy~sP#4^N+g`8X&9Q!fN_Bhz|%P;A`P>gf$L)Vt&MeY*}D+DvYaT=`OK76k6u'
    'k>RPpj%B#l1jxu)c2!)LlXZ1`pfZsU%&o%hy58U6mjN3n#I-zU1F8tOPhsrb*S%@?`wq`MI?_<Ko!(!uhlEjP4q;s{K{-CI65%iz2H&W;w=iQ5%(>mq'
    '@z!*x=QcW{2p=1XDE2Z6&h5;s(u`DvZzfrkrI_P69Lh~Y0gb1)I2@ruTRtp3JA)KLn`w51z{-T?+RlDkC@+S-pgF8VcWCYs;Xcn`Tv?1hPM|`ij)DW%'
    'rdaR#HOR%?w#_cjIBy^8`)0spFngc9>w8#zzk5cKATL(@=}AG{-`Cqj7@UYYMU>&xXj(C1uE7hK7{lb&R=#Zu-r6aHp}>%C*J7!YdbCHgI`@9#s-oBs'
    ';S@w{@H5)%uc<lf8Uz|_AFy_Pow(W}bHF%pWr-qs2+V?$NkCK<a#KHaKGw-ev+Uo%BCN`!wH^saR5FI(vh9;_0#R)6+EYQ-L)|o3R=@?1?iL#W@jVwn'
    '|Kt6@7(5Ku<un$C2-sLU)j2Ab#s!B`l_<@3!Yqn+fbMbL-(m@maWAk^<wgk`YRPas&+>>Kn(RSYR}K_-?h>u$ONwz;Qpd8Om`R`KL8dMI9WbsZdS5s7'
    'x<jE6^&j{eEW7AqYL{ypDz#ysYx4v>2*zj*tC!PCux=nK4@Idov$lDlla9_Nujc2`@I!0V{`~20A3uE#7JP}12$wI7aaD@W2JE}zlB@{9w-W;D3UcgZ'
    '_Qqt`5+_7Yy4r{sVkZiv%&SSm=hD}AFC<$ZK<TToE(tjp4ntSFFlF@qZLqexNDFb7CREhLuyH*GaOJYFG%L0uofeN0Di)6_nG%aT-DByp-g%wKW&k=J'
    ')`t1%>+e5hzkdGlr(Zs2zkSaB_1jOseE4Jbx1YXz&i?Y@%a@<{+dlmE>8nr7Qi5A^VmZa8OUe~&J1lvc6bTT@pBDdfN&|D>ZKp)uVB>#IiDFgXDPa)_'
    'iy6XD78$r0<@iqMP5eUoIi<flZu+x#R`H$H3=R1+n)+;FWSkf|M}~IWGa?;{^bFDV*uD=m?VqqYduTMkiujZQL%;g8v3n?Q>wZoNxXU9qMU24`pizf*'
    'h}7mJpL<zor2se1dNl8GGK_qR(63wkNOIY1gD+2TnvnD40P9`#JDiw*{9WXj*F;TiH}&vpL~UbRCUJq<&Qg2AoW;u-LK~v};g4k;Wj_CS86GdRQF(!x'
    'Bqk(5RfXY0L;>>GaFFZw^|~cvADnaKV<B4B_mv?t0S2y3+a~GYZ4oRbIoN579`<fH8!tEPP%=JFU1(}QWXyid3zORz*l9$3_%RvsrJ+h6BJAOeF8%rf'
    'oUT&1-4U%99me=}b4X@1m%8wsoM>P(NUjdt+2o1QH3GLyXz5cLx8s1ZqtT~s(Bj&|chO|l+f@kS?=xXfg4CVIcQU^gP8s`O8*wM{P?}qCW43?@9GgL1'
    'tv(W+q;)|DOgXNLn2YGlsq_ddb)cY;o%^}_*23sCf&Yl0U}HV6Dh1#WX852jIZAm9i|ryKs&O=}@#zT~FUb^T-%ypVv=1Qm_ktn)570KGJ{4F3GN@n?'
    '(d$o$SiBHX%^-3ptBGY!`R5ppawxn|AT5?XYc}lB$I`p6{NgP(c`ia@h{BAKJiry^VZ6%PT)f6EobsB<Tv<L4=U$DY$W;-Vh`Cf(!0h+bxa8q+(#g6A'
    'RT{PM?+?#nHSi;fPq9Q-?4%sH0PG!}QEzf|dHE=;6*+u(kw1vy;jAoM86yg$+bbRcckUL;Dsj+_R;ScI!!a8^=$8W_iYfi7KJ-%pr6Pfu60G}a&$mK{'
    '!f{%<`z0aR9EYX@!D`iK<w+MOT&C$8O%dN{rvQU*_L<CTaMHu9j)&N8vjwPo4VTL_55oS@Jis~Jb9-u8*C>WzYwDYcZ&xc7N|06KQLH*&AT%tQ;D+DX'
    'hrfOKe412LbUB_}QIV${be*KKDElI3Dky516eDD5*J>Rcf!a+k<!}MN+8nnO7=i&Qj(zQ5sF<XieY*kCw-P4NSYhG{5eiG$f(u=r!QG>5M7Dr_@diKN'
    'z=we#&DX4tw~`dE%=GPBI0`;8u<5b*EQ%0%a>p6rLjd5OaPz_u9j<1OkiY~U)FgSU5~0@}mtA%xXMq6O<h3h#2vFPiX$emk$`OA?Iv+H+#=@P{zg8pf'
    'IaffK85OQ3=VM{yVaP*C&w}r!0DISA8gP%*ch?y3H3u0mVuu<G*3UUNjEcUK#_4Okb8i5yTn{%w(v-~+>R5E1#DY0O>@19xA0G%}icrQ1`qrKygiha%'
    '{*1hc?Eq+RT_BHxXaBqKb$tELTO1^YSAzu14kvp0zx<w;dF-1+%;vTxgEI;dQNB~T3&%nVlTcKbC_-^(150&-O6ZY75vS#3LlRETJ=|{Vp<mO1o79sX'
    '6G~Tzp4uo^RUk}!*L2<cny0Q7t&CM#X{KZm{Q2%D!BDmZyXNWH+ZwKL8SNf80k>`AM4U|xEP`W6ee<OxTn|kXUe@NHw;+PFT;G<uXJ8O<O(iNixb^vn'
    '4olZcoEMQ?1-IOw)n`$GIJU!YgrQ1>=OKw2yB$PeCP*${Kz@G3n!=dCRY<?M4ZU6|Ne(6oV+dX1GF-6=SPXN;+a0%>rG#SsY~W!J*^Yt)cg<T6yRh!%'
    'y51nSz!*?xQ$b4cHqw?apiSF>DGd4#3b9n2X?%;jaeJTwI!u%_V+m54^!Xx~`$VS!uSoLtB0r;Cylzj@a7hB>v;_}ePk3Ot7z^p`AV*NHFSW~dt$66f'
    't5MH&hvAC469)N6E@C)<NoGne%c$<q#2-PD%Qi047sX`7mnkO~Q!XQJ5LX1+^*)cX5=!-;%f!2lc7=QTw%fPsQGPIO4Gf~u7zdM`OqSu!=kp69lc1h-'
    '!5>}~C8_HUFUEu5!QiriM~Zr0u7nYG+J9Du>aZdgXE+;>jb1l3E_%!963ZH{{ZNp4wcARYhj{gzn59E-<}T-Drds&%8^?xc^d{I)`aGvg)uE&z^JxyW'
    'D6|efN?5660YHv95L=J0+pUlLJIrK!PjpkD(1kDOZm%Aw!|HzhWND~j@dEo!BWM;-!?g)F4=xc`-4jarUMVY_Mc0}K6ML`?>6?C7p)J0`I<BXz;4tph'
    'rWaLS1s$rOYcQcFdo^6%RO4XMV##8+C{ZXQeI>3!=qeur5G{;dl!>w2RZS^u5v!)S>3}VY40lvG#kEJrPMx7Wdd1IUlwxp?#Fef<MX_ZbYIl}o;iSiJ'
    'eUdgz%^jSGo*RukWLI!TfBxn3Up{}OG#I8Irzupy7{itP-C5`RVPBJl(xoamMu%CaD6xUCr+kA^z6MvEB;lJr@RrBN>WOP2_oY;|tcVk=);0F7>9_?e'
    '*;_FP6tj<aaKW9&(K?)<DoaD=sONAd0kn)FRG3sY{yv97`A<_2?=n)7TYEnTLiuK63@MY4Liqv=C!6!WzxCA$(NxCD8(fkNZlh=uwE4%6ZWV#GV%9ah'
    'Ewn%33(sGz_3_aA9yR=W@?H&OCAf*AFKOaZCdPPR3QC0R3tMnbjmHH6sm2Um*dee_?(Q-4E(p|V+Vb^Tz%!i9Lc2!8=phrM0lR9%zVvH7@Nlm&ba2>V'
    'KxHCkGZvd2xG9j8=U8+Bgv!J(gDi>&sV_jfT*8;YNg3hH44n7jJ&7f>!Sbxi1gDpr`Pww}bkcdw4JI$+VC*>n#i_sOvuu@oRxyXPf9Xg0;BEXUagO8_'
    '*t2hP*&)7V4A$Nk=>+dEQ|gpip8(vRQ5HH<>N!fM3DdPfU-am`lT)awj8JwC>aKr7Fx}8$1TO_ZPx4KT-zu23&g1E^mTG^$?a~Anq5u2`m&CdK$KSI&'
    'yV}Ag-(mPI-X2(q$+*<~W?jaQw5klCi{r3Bf0zU;p9?9_^Q|%IEaZBV_lqff(bRpWmzlWHW;50}br6Oy)>$QV>@V*&-1`AU#0?r#fv+7vl!Z^hiV^~2'
    'en)slgzmn$PcGtd>DN)I!vvJ&o!e&3<(T-ZWP;|B!@0U&34~XCB41@p6e6)F5cyLOl->C%8v`uvgo_Jh1bc>%?2RnD!u01?nWGwY2*^W34YGtH;80LF'
    'g%n?Hk4xAHzMUNp?FLkz9R890`ui7b;hCg)VG|1x6ZKj651`0a9zv0#p8G74<zjbSuQGK*UAk$jbC(P>&y>1IT0|AvTj%{xMh~_o8X`^vw$z51)4IAg'
    '*-@%%>HH9;Ld6u@(-s`AMhR&qE=0N$;J*gylIEiLAJ-4&8pJ$jcvhur4V8e#{#vDd1~jAs^STJMxnXITP+hz#E)9&gs0VVzhZzP2Au*p*OW&t5uT--0'
    'NrR4MwF>d^*AxB&ChQx`l)NsyZLgHr`K};ydJG~F5vd<Gb;fOOwjd6mngY^Z_jrvYZeO6O1U8!r5iP=UZPFCO+iDhgDu=O}lWD`A66|bZuk+w=fv-ee'
    'r-}_WiV^Zq@8evbPhLng8>v!xWzk&h<#7;>A`hZ}f>^PHo3?@jI_<v_dWEhQAog3&cPxwdDMr8QM%8c*FhyQm$7q+az+dH2q~kI_2^Kug7J&rLrD4w@'
    'd6SN`RkA%Dl4pj0{=+@aS%|oRZw>5ls*C*poW0AE9LIJo_$y%=v;h>W5dMnCs4X%pDl34@j6C@Ot6&;5!~;7<i6CSop)Q@7hCS$k?0!d^-L{_Hq4;CH'
    'AO4QWtR!^LJxY<aSNPkvyYIc0Cy*CPOp3-=6+FSAsSj5b((V^jNMlrv>B#(JSnr$7q4~}PG;<XirtQJ-p?CB-MB%&E(39$mh+e^Yg5@ttVRx#lYp?=0'
    '-CWTs5h{CC(X8h<gQdz)%~o84mGG2JtLn0T0I#mE!Slp>6DTL43{cfo)g+&E4#`n{BT{`G(qXo_JZlkF@2YXe><u1jwc2$X;Z;4PK92!YvBLxs9QNBl'
    '_ZOz64CA2D%{d@k2ua~}Os6-TAuE)skoZ95*48H4irhkhF=fF~db=nJG>K%b@IJxj)zOyNhK6BxVEM2gH$CRx`(~lk_LyHtz!5K1USm!0=w}ScFGh}e'
    '&J{IK7d^R;;?jBItz1_!(&;7w6U@mE!px!W8x$``g5i5l4XMzTZZIT_oNVF(QZ8JaODv=>G5@q+kw4B&fDi7=&1>f!VOHen2wR5JK7Dc=C&VgV$tG!+'
    '#_O;0&*@k2HDqQtOwaLk89(*Q(ArYZmR-|x*l}3xd)uCElL(ng6(H0+1Qje#g<vPpv`H;R<KO;7wBb*{rnTf#Lw|4*uqGK*RwOL^JcGOIexQvJUwbeH'
    '%))7*6CqfpxQc$BgX5mI;_-2ah429Ua|QtHnJ9fX(n*m@V#+qD7++pSKj(ogQ@-UGtIG^S1`K%?W80rL-I23eSR2`)*+|dTZc~@;DXXfgGXn}I6E}<t'
    '<tCo9I!yc71|bCyZdU3NI731~7H|2}54CiUjk?lecIYIk9xNKFLe4Mqk2QgcY_|07Bz73c--jt}a7st_SPrQTNbMC!7RdATkkoVIbdXVAT&ibx01={+'
    'S73&9hp_EQXM-W4k4T<Ip!Ng!vC@B_WaiO7`2jeWd8HnD^JpSK*?bdV)UpDz;H$Wd>YE_r8yc#&jEmWR6DGH9fD|!g^-YlMzIn(dZcMGe36jXEz{|09'
    'hf{pr#;Ed}P&o?378DN7I)YT?w?U?gfz+fitgOEYHeBpei`H>LziqtJb#_3ETT^{oYx<{hL7Q#dJ?0w$4_%m&3E;PNPSAJaB`QYzwxJ%sb6Bs!_&rIQ'
    'YZ_A0&JZN%aXCxelqvvL*8_U8J)(hw<tN2#17>?QrsL+f#gr*z86eB6AsySEg-o0kX@Y}hnyVrGTH#QeCM~E51vOWr#?C);Ks7ooY^cAvzW(g2yXt3%'
    'RY0t+MzmC&w(50?r~{(D8j<NmEnu6rWgXhG{u6*sz8V*b*V#Ey`22J@C(aM}Q+|C}l`&5UT`N^yOulNX*~;4COy{GmfWuM|(Q7eCEuy=r=0y-S52Ad9'
    '+FHH{V&*}7>`v9z)r%nYJV@=YJNR+xGDsVNxMPkC$DG-!_-k*St!vV|+YYJ>P~{a+_QJ}7iDdQR=ho=wOJK?J2t`ZgiA~53KVL$N7e)yMNQ)LuhVvbl'
    'fyGNA)U!+*O@~I)SAa$93+l_JwG0I_*Ay&TS)h_>i`AiA^)=-RYYH@e)?Q&;)a6xB2j5y{n?i-rTFYypm<ZLjess!cVt+xZi*VN*Gs?{2{6@s&oFR=c'
    'F{b`D%>1CyBnJ)gZJ^mnqlr%%s`@s})Gw~%lZL9l4Kj9&@diAGY`zUL@{FP0aR*8Bt%!%nd5@33Qhz&Qc=*GIK9h7%`)#nP%iY9>U@5;5GV-~p*ORl$'
    'H#BH^WVN1+4DOsPufdd)O{&aYL=_ym>jwOcD~~fg1q}sNf1;qNFW;oTyb6BebxX*cGhe5!q728F>w131L2lCXfegizf1;R)hn#NK%z+yHyb?0n&S1$%'
    '5yK%yHh%{35WM)T$3si1Khctwn41a-lRI?<!uCCXFqnDhCvA^I+(Aa`GV|t5<>?=*W`4D^;qz}h_BtsB2uVt^&>x0Mh#MPIyVNmN&ij!3mfvl=ZI{_r'
    'KE+P4gw8ozXiTW+={NyseNvd7-)tCQ0SK%y<ph~CqEQ>W#wF1gfS?>l7#64fkTF;`wNJWJEC+oCV%hKC_aE|_)U65D7n;I9^JeyqlA#+eeSPHI5aNEx'
    '3(<HUT~*FIYYcMd6-t9`E9%e{)p=Xw{`A=%c^kHM8(N_`4|2aoIoDNY-j;0}dhx<ogwOYwH`gtPEGb*GLdZ}La}k59aL56Cv>x=tg+@G3#7Jf=quEql'
    '=tj?&o7pQ7da%9l@gE*AQ6qlGseRtAuIE2Yfit2CFS4)beed;|AR4=L*2Z5is0tRKQfSk9y)3Oc0w3ehi_NFsKmPpF-#`BKOL70}$KO5`-7i1=`=?)j'
    '`}pfmUy4t^6w9x_{`S+qe=P34{0e_#Luynl%fPDM&h}VHzG9vZ|8le1V6uX<ZsvEeV2zoxU_3q4b1W~%tk0k0_<0y79={ON+wi^{wu|ux3cqi4E0R~j'
    '@MJrG{<TWMZ~KhnEIf0SxI9LiR_h^Udl~ZJw(0R*kJ%Kv;a#_Uu#e(Zqte^aCfsB;?Nzuzq&7>tEYpLi)WH{+zaEaDn=(KnOMu-FW#(8l;l8Y?Uw}Le'
    '>l5X-a<Vmj0D~(=mnACg*sdyAIA9XimA&(lKE-^BKlDTsMeao{U2mDX&B8UbHZBUKZRUp+PqeUUbI5Kn*|}YqsCzJto}g#fNfu;IUQ7k<ox3EepbSw3'
    '^C(zDQvFyB`*j2dPD2|bg~g|wUwjB@(``D61LU5zjh2O;Dd93tpwZ^X41rOb(OfQm`T7~TdCDBfM~Jn}{F_+<^yBAZ%RrM7`KX)r%x&r!Ae{p>CIFRx'
    '0;tM>WElq5{E7MuP?PO9IygOI{yyf=AMFOwP4@d%2d_uWFTTOtQoJXD8FdKEn4QXQ{=xt*9*erBQx_&Gk^>jw1MNzoOd9uL0=)SKc+@{l1lo*{>1nhM'
    'D{?I^eIZ_eIBjrr5jEKvuTf!nQ&&SioOVbclN@u6$k+K$+D80Y=x~?<-1H0?X1}bjFEZzRP^g)UEy@}<JTlehTG+P(rOY|R^vSLXh|SfA>wa`wk!aJm'
    'pgL?tb$u<~2KA_2`O;Hq9X8^+y%sNN!L(y^INZ-KfH@xBlq~kXquZo>TKf9x`A2wGyZ1f;I}6q`&?bMS(v^Q+YUUp&;n=v@oqMNR{qOug&K0eq*Cp1G'
    'bm=7Ue=T|@tt$Lhy4syKowtqNna|m(q97A8re+G(`bBHQEo-gE<rs{s7maIh3DD*amOEa9R}_^OFkY^D%Yl5Ei(<UzD`XUVm1h<C<gxijK5%pz4->h&'
    '=V4}_`A|V2z^E%=%sh<CwL}<m1x!5;<Fzw|j1p<i!?bfS_GSZN;7nJZg$WNQ^KGoO<2csI+3*66n#XbU#)zY@!7=kV?sXV(s%vm&o7s43Rh^vrFW}6!'
    'bHnl~7=EZ#<c#dP{qYIL{-)==g*J-0x7(%7SwQLshDm{2*MSIZ&I5!09?qJGmg2~=^^e$9<k`BgZK|LNrL$4pKKrD9-yv1pY3(WVc-EtW4g5RDma%-Z'
    'V5f-QjW;({TfC~uXI;Mb6oCvi+@F>niUVWMLD^gB&5aRfu;93{31I3H7#!_BVEpCgrZ(q_w(VB|&0IpWd<-|LJcqR(-v+Gu5*A%SLnZYYBoqb?Q0*ls'
    'uOSQ_pw5*HpJ6;+eX$2zY$#N8+MhL%)FIDn>1m!>4`zag{1=QE=M%DX$Ebw(w`Ul!E`_DylrzFQ?z>gT9OI9ucLBZ!QwW$m1D9Aws@C$?B9Dg6`WZJ&'
    ')iADLd@lQak1b1?nU5;U0494Q<8(J8r}pEPjGr4MT>(#?O!#5Tf=JUcCp>MHwVN~%h?5*D@K3#i;Ote|2p3=63LCSL*dvh&$Db2vQ-uP4&*I(RUKeN^'
    'SBS8@n3FE+^H|US`nTJOJ=Crir56yD)3Oi)#_DwFxLq8vlPF;w`P;7r+J|MSsY;}=vCC2meuRhspH7?2lXWBr%sX~4aA4Pa!Z2d0tHQ=78j9AGs!cy&'
    'YUT0_TR)kj54hI13~SHE^C(AU)UE@SD%A`G_|&dYDoG<pbETb^gH^02=9_WP>?1SP`snMWvwCSu3xT<hOp27Ejt6QqqSOl57BsJ)tV8|uIjHkB)RotN'
    'P%VdiBmlN;&-WYbNviTmzZ!``YvG8YMb+?uaoi_aoNoVoa#<`$Ll0&*oz)-J0?xNh>jijFL)1p35Y%DV3~M?Hq5@vGTMvIZk#6K>IBl6)3-5?hN}N{u'
    'X96Tuw_wOd(=GbqZI1%)DUXeE-YN3f#0H)QhjTS5&tXGk6)}sAYBFA)w)8@DEFfty-w(v*6vah@jbXteuiyy<X`^y{fFtWMVh>ZsMXI@<{D%wsXh7sY'
    '8)P<gZ(_$2rg)?rjm9e9(Nl=stq5h@EQW2$eqk>`8)Om1$Q-XwB-ulp7`oJ0A&a{e8sr~WrGE+0Q76S1_HHFNl;I(&7PE#iy^6EaL1~bmz9v>7jv^XL'
    'UgZwPS(BLk@qY&*vMBDxUw)}d2WBSIp0&kqXJ6G#_-ZB2T81#1FzhP-I)z#J!Ia)ba<WR??TKq29~N?6sld=!GM~&G(32JK%gM1C*STfOw^Xt5x;x-?'
    '0>5gbYn<1i^s<y~Ycgv)y0#Gs=^>s*!bqQ-ETs*ZjqD$Kyv)F|6BY1{kV?s{{q4mgFCaMpW|G|x%e=jesza5V%!aGVO}z~h8D)Z-pctH6c@E^1bIA6O'
    'GNFHROZ-yM>WEg+yr1eS6k6sWP}Tz+VsQ5jL%S_uv9@6$xbky+0NX9Q&2Z=wzA~f%K;`Gj5aG-hyFS(t#BMnL@t@oLuPd)1TdGjd<->U2$9oBeJ2$oS'
    '2(7@q{oC<}#pq;IaRw}hk!V8K-3M4Pp|*!Zpc6*b5ZG#cDyY`&1Zg}i*X%O0wcD8NE~y$Nz@>rJM>v5=OYah+@Itxo?gz_~Htz36+S@i8Fut**T}UR1'
    'yIEO9S~j|Ea=2BuIoT|>HFn(FQ*V=BDVBtYwBMGsmH#AK%NF<F)@S@ysugEYc^F$ZJ21JiI!mjisy(ZJ6*&VQzz^kU--D$Jb=QRgZ^E4cv`{*?Jv{Vq'
    '7QMG$=Y~=G#<JiBAEK!oKapvs{T_9=?XCA6Ny|F&DeFx%+%j+)H-@y2rAs<TqK`@8Jl~leSAg;ZnB2$&MLPC#LaD@A6nl_%D9Y`-3PW??s6&S~Q#H*w'
    '1WpKG?opK1x$9cH^{)0yZd_vJGA5dHsEx=Zt&!Tdl3wkbfASxryz2_ut3#*tZNESCPu>U#Hyx9({b@Bptwj*CLdW3MdRN-F+k4)8UB0*p)8ziNeHx;@'
    '84oB>_!>rkqr=s1ScEJGVYfc=G`R2AyKZ&~ZB$rPL}*bsrrw=cgZNFTYw*5A>m{7*Pj9!y_e;#_j_-DO0TaO|m7#W&$?w+Xh8<EEBQiV$ayq*G>;4fo'
    'o-K{CJ?4`;ugp{eU)Y=|XJwPE*=R)0>!bU~3dVTl6?s}whL+nAF+jIpgNqDKkt%j4R6Go;dEmruk!RFMchZD^7Q<kBkDlm}X*rP=yTGH=`+;=+cpbUl'
    'iG<FY`=tt7n-mf5Q->5#C@Zsv-`3&6kP``uW7P<MyCIDVBX&1HF)WSIbk4(KQZ+ZJ>=>3o1{Xlo5f+IZ<U6WgoegWx&~*lL>z%W+D=~Cl=6sZ%<O10+'
    '1A<#YkHdakbJ<rT(nIs)#<udDdsP{Vyzh}h_88oWjEcJn6liVah#ydKDr*#{GfjUg(m%DE-zH^XM=E%Kw%1{d9N-A$bdB~&Dmz!_0G^y*t3>2IyUoBZ'
    'dtq9WKq#y^?h4iZq)ExtiJNXUEd7}j+Rq9it(Il!B@A1RU7DJC0Ue?Y_Vl;V@TDhAY5CAMF}ACEX1R888f>ZNERn`5x7#>vGq%%nb_qREV!cHY++c0y'
    'tlgt_UR7xAa61LYvkmN=gv4}bNzbfdJ!i<VE;G$*f?+dfNY}D6c2Og>^LB^H_l{**hUVeRfZWdC-Db9@GVm<5%y2)k<p^PS#XSQc0)RXZFdmktX)EvF'
    '4`bXxVs=cqtK!V~Zr4^K*w$fO?e^nfmC!rB$IY3aJqNK8%6HM7^E;sPYwt_Q*8Ks6Pg4;bmF{Y9%5vq%dxbRdUmE(uh|3PPO3aaRG=IuZ+_(NZdq1M|'
    'lH<3*jYXpxv923eU|6rInNK7JfhPSOoCx--F`}1>kLau(f_J}OJ%rDga|l_gdN8lN+J&}HD|6;^7a;KWtq!*8ob7~F9m=dkFf`{3(`yhCENPp{?-4AR'
    'ms;-o<9?j-^jh1x3AXJyTbHvH*){$az@~@H*hQIHQSY`Y^sYE#nNe}tkG5qftdwU=Q(w{=im0^ejA7g-&Q8i;r_b6&bivlnUoN>UiW$2&jmI=dwQWoz'
    'j9X(wBSlnH>VKxT$^0a{bI@Cq<d>viwjGWbb-*l4men54%XmJOHV%yf|Kg$Qhnw9kcHN#Ch>&W5ojQr}-mow8=Fd5dD)a~?M_4f#0AsFp!g;$uqZ^S4'
    '(l8gI(YlT>@G!GoVvManT=J{!0nSsXb^)IB8#2x+Cfv^5Xv;+u5V1V=Zi}i|JY%s|ro4fzY+0uriBX_l3y4V^P6Hg})^|3gBHqfu|9Y0?vk>2#Vwo!L'
    'lvP{$3n=@%?SYtIaUfM03Y!^r(P3gS9*H3<r3tmXek4@usyM;(PU$_kbIE#FXyOSK#Q>X5zs<_()E=p<D#2c|3RqmkF0gpzbrw3{n3r*)%Sx~4Q3sbD'
    'iD)Mvbk7;y?K5EyunwI5s&2jTTy5`q0@k=@#Uab@@6-x5{T+gDna16Eq)QYk@G|TT`!5lcmG!aI?@j=3MP^<Cx^B;yIo?gCOfywJB#j6|GONUCx9r~a'
    'hkR6u=2@8Ug&UOy^=nOS+7Qwz>^)e(LT=d860oeOU$@+&@K<MD98_Z}P>u;2dX3!d$$S=76Eb@EX5It)+$<Zn;R@J9lg3unQe<&e%LIi9Q$sbkg<zx_'
    '{xd}~GD1=>AjykmGXCYr<}d~%^x+PFrd0k`TG$Z3i&%=WD)^7X!gE!l+#%8b3^*<*4+VImEpGMDwk|IXS+~+AiRP0d-m)%R?DMRa3-ClMuRoZtLn#VP'
    'RExszje#+?<di&jL4teB8Kj^PR@uRKzeCev0;jyja;f<8d<Je~uWMe(?=^~xy(TzSIylUgk>JNlSQ7(tcD4+yVMp-{+V=1$0tt?Da^gSdIdv#13R&(p'
    '$Uw=nZK7Va!!u@=J(wxhXta4!>HJ!+Pw$7$9SFc5c$Ov?fK#SwgMB+^@0obdAvEEIFq@keA%uM85ODq{CTjZ{2BS!4`-7)K4$9M|xB(03t*^IxSRV&E'
    'cO=SfLfVOd)b$q-r+l|8Tdb5eH&!em+JL<HdTFzZ0_M0GF77^cEGq?TqHQz(Ge8}Ope^GzfafTefffwZ<nUt1w01mnNC}RfOoue|A-`KFSn_}L>pk9F'
    'Y8WXlco6(eIwG;X7cRB)(-@NIfwiltVPK9+3chXA5n!bU>A=Eo=nTg4GF>WPRN0+|G|OnXWRq|Y9twCV;Mh$pIlr5PkLXS#vZR<9WZ$O%q+I!k0Pwoj'
    '1F+`X8GuJR#Kv+6sKo>S<)Tl(A*Yg~KYbVus3YbcOFPstTNm>4XBh0Ce@8vQuLYPE$MJ4-o_Elq(Jlj|F}&|S5Z~)PUe)1U6ov)(E3Nx}eFNLU5kTwc'
    'h|Lo4%Ik;LB?E1s6|E<;clv9r78B%(Un!W6<==kg>;2)blETX*E6O4pNO_Fz|Dc_A<NE#Rl}cvb)3woFH51M;*iFOseodm|-}&S_tloFUO#!Pd`VL<I'
    '?FQ+xKmGN$43d=+7?SSP0;GL{{o{Wu!Gk@_W9nH<-2}`g!Ss9Grdz*<l@C$H_cBA3bEvwFN@@~RdgndT`!c@j&lj~K&|<$@<iI|$ShIR+4D7Z-9yXgl'
    'RbZcM3<XvRu0w+@2+moIX|Cz6puenZ^85oPd3$>lvF(-^279R$a@10Uah0ON40KBO=SNmL3fwwc`m1uLD)yj@?vMZK_xW3j^1g{I(@{eqN83}6d$+tL'
    'C&Im)*ho6>l30JydN|1x@ojETYBmBy-l`%3L_xkTRX&xahm>hQjyTb6NC)MR_V@(pw<vLIEb`qFQ4(aTvvgWMSm$I+hbJVb$?6K2hgrFsd)%*DkAaH{'
    'D&eAQdcg1l-i!IX)6v62{$4>({$*9xm1lo8O^k;x7K8-jeK%km4%;W%Pcm&7CELd5nDf>!19CX+DEMdmp@1*{0HD8p{Q0N9#*u%h>k?gK*L*qbu^g-S'
    'e!TBds<b07<JG_Y`swpeUq8Q2kIJI-H;(quGi*1wa`Uy@a420s2AmTl-RBO6aYL`D9Kg%h#q`&3&MlexSb85=31UZmOA=<d-QyZwi1m?z{^(DgR1Dr3'
    'QyhmE=QY$ynXT8QhR>Ftf41Tj7^@4u7cAGtOEuC%@LQ7Y?C6<U6+6_LIl^GVGb@~uDdxK0g5R^xfo#|S0T~Zu{}>}D*uE=XVX`8uGI-0`??-Oocf(Qn'
    'dogNb_$P&O|E>!EuC@QJ%D5?76=dQm0-FA)fSqDHWS1)JOr||+RPC=GHBBS5rz;wmng56vog>VC{6$)V7)3VKGcwAL*^()eFQY`7pH)i>5QYW{?;P_r'
    '|GS@lFaG)o{`2FPUqAl%(;xreFU1}FE*4*&iyy!IQV3OBjXKQ#|NYA^@MHhww@=0D%g=wkwKsva!t^f`l6G-DhM&5zgkhF*j3%6Ta3#Xte-;@)t(|nM'
    'skE5^a0&OD@o3W{Ob&*1)#&n-+QqC?9FcSmjC0;R5S;{`ZcI?nF=Y#jwQRiFisqMG1mwdt3<Nmb9w;%LU86TXsYpciMPLNV4yraknGULTiac|{<?!;>'
    '_j2t<&{anSLSakW`yL-#Lcz+j9ig2{8e+-wShz`bOX$@%%W<*qr-&^{R%1Am>kUy?O^S#wUlB(pCp50d_dU@vI7DbUL~$*rOM~(EA*2GyEnbENpZG7Q'
    '0blp?f#Y`=!wCko!x+C_IPIG81M~o_)w|^Z_B;IJc#2^TwsJ8Sf<uaebu@m5<-7hVE(6vNjq?fXG~onlZ>DIxJ#q+*kCr$|4eg}7=$C3xd_+2#br{Je'
    '85@T<J)m3kC}TFmClZ(R5y!9!lf#G>r=0d)h;Gh)0bihY<Dq+lC4Q6N1IAxDoNfTrNmd9Wgg_Ey7uO8KRHt9D{T+?Y^o?X@aQebGUOS%C4RcuAzE5<('
    'PF653XA_NfQ(_>#Pb@H|nE`w|KH=8+E*}C>3_oL`YMU2ml79EGSWu?ZmV{_SEdOLs*-$5%h{!(SQby^Dgj_P*1y;Pu-+%h~r@t4=ub=<z=T9I1_<!M_'
    'EPnp@3+xF0Q7pgw`sJ68MXic|`{~o)tseyC#0xQS#VYS<z^>h<<#ZWxS#*yoaT-B6yL$r2rjpf(r9_!VOc3o{;>)4!(7(qSV2aZ(;H3TalP{hWbTh;}'
    'H#`MIc@FHMM?VDq)$z$n;_i>6$epz+0N(gscf&}!Ya(Hls^&itlf{H7kd6I9I$>mlN8Ok*OffjNHs|2s$lfpa-Tg6H-?7XGEHe-N1_k|$QP4OpCKEPa'
    '-Ch8N2FA?G0BF=rDeLztYiA>?u+*q?fRR+kOv~DUGq-!-?X8;wEysP$g6`q;K)PE=ORC19VHBEa+NLdiC~Z@o9ak_vkIRR|FL0|7?`aLIB2b4k`Lf|w'
    'XLM5D*s1A#f;To6wNYu^5i;G3TogGScj|in8`I^Q@o(t*RfszJ3hY^cc@v4XRg%BsRxByCx%NBk>)h~T>ima7h4B3x&m+pK!d+h>KTQ_>u<DCO&+Oie'
    'FJz=TC%+NDZSt+c^?Bo5RP1v-4{xsm-&jwOUd!J2m?S!FHGKY|K5uk#>@Hnq2d<>`WD41DP%M`HLQ%DO!;9<+ML5kQyn1!hJh+sUj$rAlp5nv}O2N2O'
    'Ap(>xRN?$V&yE$p(!kPAaV0K;(WNYsJUuolX9QwyCt9ZIkisTaQ<56pp`C-<439Rsa*|?*qK>ecd0BCe_*r8R7YlOn{fE;XhSTA;+jDLao=6^rCD3@K'
    '%H<~LT2^hX-5am{$(a7+d%(-u^d3|-A=E@miV>qg|Kff?jZJ&QfM<>_4hm1NP3k7RR_Zcf`1gG4*ydIyBvH|Bky-TzjB{P*Rzihui8(8t9J!f4*4jMS'
    'GiDSEe8`QfCKy*!V}1bPIO0R`^QV#z<?GE3HYy`R9Bu1t@s`V9m_wi6VG{0f=A}m@SNz7{;U6EfTuPD&?VVs!V(zY6Z(tU0T>_mCLE|C$TMp|q&=FSs'
    '30OSaTF62-b@fK&$W9~{k<InUE)62at}^V(s}BHnb3L|mLC>^>)0HP7;_Fdeu<IFAmokB<<@Kl(lbW=G$dGjzmACpc$cHXPZt^m_RRxj7^~fA|G2aea'
    'MB}0U6PPyGG%0YKg3DI<0<*~26#UG@NPs7>x(TrAI@ok#G-cKd9K()lgQ!(~735?_gou9&8iLmGF6b)IY$_n?bc6{niXr*==w;k&G8B{@khQ80K9_uv'
    '{4#1j8S1>CRRmdG37O1-M(3rJIzWo6ATJGnQAcps^?Nq*V?bRRcj_*6RleLce9M$JW80X}<NWHR+7P*=!$>-~tn<{7sI-r;#T>fu*NY+BeEuVjvkmNi'
    't<N4s+1X%!M-Mo#6$bH5m$Lj^=;CX9f{w1i1G%r(at1B13|jG0>z>{ei!VR__7Sna{*o8bM-WI=w`V~ub(kyV={<l|^95C9cydxIEZh-haC}42s!%l7'
    'SCNY%$^i2)?7-VjzG>>x_>g89)*WzRhxNmFI;4M&c_nqg*5S{0!y}aMcz6Fp=2M1{X)S{zaezU*J$5S_9?MSFx)-8zg(=(UkALW%itq9$aPA1xD(7%Y'
    '=DVM>fzOQF@A5SIGHNH8S@O6%d}m7r|Lp!0%9j5%=iz*p7fsipXeKjp3*t=4Oe?BjSY?K^K+Rf76|Cyair1j28BA7hrkS`jGIIz2+{9N5Oq$dr)uLDr'
    'o88duhFxa^QL2g#2;rY|@MseD31^vG|D=ar0kY$|+i9cd_96m^sJhT}nu-U0^9f;Y5QmF(=t^c*Lvd5d3=C_G=;hZ_OP%tV`P>Tc@5{)ttf}$lEZed!'
    'Rv8bMX+4peBF_YAM+PD*UGvi`9rLWV^8I@3>?3FS`RC#lXP(}{5uKiyUl$CREOw^!fy=OlP)n~lKU)a-*vQAWSbzO{@#~j=`3Q^eZ(m>``clZQD}MP>'
    '{PfF@Kl^s)0H0EoURhO<bL~R(U`c`*e8Nb;#{xcEEFJwK<#ob&Wmr8NIR$t4@F7G{(T)CkJi;i!e?ie>@yS|ZQgHZlb~D4DKf<1jB<6SBU3h7%CEg8d'
    'W91qo1yQ2iX}f#~!JD>jct)icEdR74I3)^$3l^vS0s{|ZX`>qDZB20k@p+8n8Yu$dhZ`6VmJrS;8XOYZzVTPebOVUU6l*!|iM#;*5MRik7R9RuGm)`J'
    '`@`!(J-e_88_Gx73Ka~|+R^NkeD#Ttn{K;0Vf|u(y8`;9hZ(SkO`?CaKVJ*11i!Ik6iEf=Fe%)9A~XQ|3Wzx@F%9&NhCVj|*}C#(l^5mGvl2ibK+|h0'
    'KtjETKPfqEjgJPf{LqguEo$6V>Y#M4sktbM>HTm3?{;yzzkfP-Pvw43S5XpOmD>W}T@xr6rtUEj@4RNA=rjTU_>SqjBzrx`jF)4XCScXGSjQiR!;whC'
    'Cve7OWK>FdinNwC-3V?qZdXq+4OD|J!tJw=0Cr0}Nl5%tW$aQF#{#Lkw7H^F$n9x?q~R3GPixsBBY<*`{3tU|6Jm-1z7}2{*}6fqHzuYkRG8i1FR_rq'
    'sZC+1m2i~97~0-Fh5I#CQVj`evk>4Q*nxoYu!2CgtFBnwE?~hq;F|^o@ofM5f0ekbr}0FJ6Ug`(Lt5+65kzsG_!}JINv=3Y5rJ*7%mS5bql4>GDTlWi'
    'x8Qu9aF@f383nEK<7Zq{zjw>&e-xj;{QFOzfBN*x?_X{U)xl(WWtZ2l7>8qt-xWpalDM`@uWl+@hv+-J6Tp-{3>f?rTt6uBafQb@w`k%3MV_+OGtFu4'
    '09XfLO(-ze9-EEhCVz`m=_!U&+Vgj7X}8F39i+`Mo1()V>OIs$++w9~o07)l@^(zoDytNPVLN;SW6>7l0x$7KVzhn%M48H+VGL7!dXVLu^=%aO3j1><'
    'Wgw%d<1(CJ|DbK@sqgleia|JK>vOkle!9hhad^07EtIanE%nr2FgNPR+U3gNdisG0hhSE8>yAUYLy)Y1!zcEOkJ5635Ml22-3locVOSJI$RR3W%7je^'
    'v$-7stf}J~+z-&5xDOG1=$@i(SEqGnuPI)M@3C&*V}RVRuXD`SClc3XeGR7P2EPO=RTE&%HLxt(G{xk2b=f-VRQq!l?>N#-qw7`H9OxH4a)}iG8@+d@'
    'g%tk&tExD*DhG|r#|1`EcYiC(M?ne%X&*POzXPbX5D_d#ZNS)i7U|e*OLr$(Yfdk+CNbf`ixW*4;oh9=!-zvkbU`whDNd<fV?#~qr~>GTY!9fo1KlrQ'
    'E|2hUY?@P@EDpbGUofY3G^M|^t~rFony=Xbd`oUCv?#?|6j7Vf*QNF>AC{8Mn)~#N7nh7SWHV_dQuh#t2YjJ|pW|UXhhxdk8=u;$o46R^m;*aLj3k8|'
    '=W(VI47GkkakF=pNF5XJ6C(8>LCQa`#oiAam+o+Y0lx1Yep#jXRlu*V$M-?RPa$o~eZVr`BNGg!M8KjZ{6`uoXFoEdc=l#F?3X8eowFrKv~jiZ4E_9b'
    'in$s%%MwEi!I`RcuEYnOpC~G5(+14;ddx_a4It|jG7#aL`g+JnRtX^W>ult=p4k=JC9XY(dtF)pXy(DRmwp9Hnvdfz(iMzNq}Mj7nbY5zL=hIt@iv;J'
    '{35K|vJFA?(e&^yaNl+N1sGGt9|?baSHP<K<7e>I$}AO3>!yy->NzyZvAhNhN#T5szm|Fi?0y`Nv||$n#y!uVh(Md8<KZ2c&d>ve<sH0GI8$AKLJm(4'
    'rPK`64n2>Hr75eZPGSy%35cc`aL}$%x)-0~oKvWqa;`@4ws}7e>x3dYZ1Jzt)z%2laac38yvH3j{9hNnG>O-fF6u(mOeJC=k`hf86Lhy!jrE23@+0O='
    'AySgiSJt)i8FQkHM#G`Ihuy^H?U4_aZi1~(+G?wKjplIQm-bLkY6JUW?4rOC?BSg8W~-6B*E!AFxS>LaG~ziA%#s)?<(wd<b=h=(IpJTvMUUY)*hh5q'
    'P4}^3rQGSx{frFb7p>StdyO^zIWbAZU-|D;Fi~?R5mTTmSQ84?%$fM)WA}|Ze50N-;Oo{U44>34d`QoJC??jhX9@5$m1DW7RGHg_O%$)+Zl87W>Ypj{'
    '@#dFLfBW^z&97hn1%7_~_2V!9^6|B~;aiDKhiC03aQ1O`>`~DrP;9#F|IKpu$T`7nxYMD(b$6`0eShE4boNG`N~aQDhN9G^ijB#mSiy=mpx-ja9D#1E'
    'xS!UB=r|^y(C)zrKB3y^6BZawua8g9Rc2H>841*X@Rh=#Yr4sGw$aWdDg`TH<zH1RVS8i0q}y9K+rdd^z{7TQBllbto8`^jeSt>^8?*3_OR?Bk#S+@i'
    '>sH1{c<AntP15DH3fRSOU{S-ry<yV=<}>Q`)M-YJ)Rws(FK^q5XU+ad&maKw_SAo%zh6A~w?d)$?O@Z9SQ>IULGdx*Xm~$#MDCf<4pvrSj<Uq<K_4$@'
    'y5PTsBtMt9^|;=YGNTUCo+?qZpklz;<0I#Hn{hLKm%eaSwSV&|g&n$ykuSpz7VUS#HZ?b;Cmj;%oH^w#4BLB5JzLw5+TCgS(B)Xhv#h9k4hvbP#!CQo'
    '$TPDZNCI#-bc?lrlDvT9-(-Zcn_+^}V~X@}%xHp)yCR^bET?7<!w!yvJ0BL2f3=mtTwE}p&@kBq4Qr~;n-a~-G5@d#v`E74SF_f0Ojmf_;Jq4AT_|?g'
    '`=T%V*;NW%pSML-<lHonKrze08Pgf2EF%IhlUsG>ZL=N3&KN&TOJQcL-*$&_o7?DUXv|b}ciOKe4O@j$Rx_jU^sYF(JFT;^pi1MT291dOgDxMB{UcpR'
    'j_->-YE$krqRBdnIaG4a+A7~=g>9P<)R4^G$PP2fE&bKQ&`%&5pWG&NUCu&0fS2x>Z1L}+uQs!Fy+i5LVYGAvIjyD%?C-qoP3%#i7-r0_dkgLd9(}=g'
    '#0#DMqyhX5w_sk=3Do_*zq@<N*6gwg7vo~K1hCR<mk;^Z$`E&1&sH2-c)!5R=ViCuc8ehkzSP>UHZV+@IfS<(oLKkyhXZxJnI9kP0nb7G<G5cUqiJqm'
    '>tH`ya5yMGba>u<>i0Pxxd_A;M$V41&FL-p8>kVSn>InET&p=#pU<2p<%aBK<`0^&B+^^tWE+n8cm1gt+L>-H{`k*e{>Oht_M2`xC!$!Ps-68djt*pu'
    'fqgcU!3OidEadJIGIMZ?j(mtBi9~X0?z+4MeDXbX+HBh$`Ym``OnB*{<x7m0R~5epTgsfc$8D``lDqf4!*XVz%#~9kmz2pCneT_yK>TYXE@!8M7fK+;'
    '5?ns{5puJBzv1#@iZZjgT8x7&y^V6%jCZgqrP@bQTGy6d(A`FZao-Dk&;gg^!xK2sP5K__C^uD5R#p7n559*bJ~<-kxUvBdf>jtcJrTYgPyL=8dyE&y'
    '=!tgn0XX^itw{3=8!<dtLYnLy?dzw1|5W_*uOI*Z<+tw(%-{O?<8Q^}J*}wZGt~82rp_g@X#0p5e$1gy1zw)bNHJe@-&TicSIdMLs5DVKCbB+3Wr+jH'
    ';Z)6!pL1i~2IFeZn0$+QTNS7qd3~MQj!s7f2}*|jFb^aHcO;*1&I7r4q!}2~Mlj+s3_W<v6{U=bKILT$K0L=DJakIo{eI0{m=h49?+%UoG}}dXbd;X?'
    'M2RZuq9sg}<KDi_+1u`DbhxplLdt@P##X<+ANjJd$H6r<lp(l_x!bLwEwba;#AN!z+;{76xQj=nH9s??U6uXFQq1sbjj5LSAl}V>(qALWHXh5z?sy!h'
    'pV5^FbSlQU&K{!^A-&{+Pxf|9rD6^kck}btA3yz8{QT*!#Velh(M(4VIz(yI;EO1C^{f>O_-y)pk6|$dCN2#hI)9ws4*n@J5A3p7&auHj6kc;LU?muv'
    'RV(X=ADLPSBWYH!Mf68}JW7a>=#mBFAX2VmfN5o54Cns60AoYIJz4@*<7+%gK9$yiG02|9MS}`->wD+`tj+ozuD_?lp~vC~y}~e}%7<CO=G6YJcx@Rv'
    'U>9ge#l$L4V4PXW1DTo_2X3)OHuWR8bEwmjMZ&03xb35Cd21%=I1V<kaVnNNx@4Yy+Z@{e?H&$=-Qf%|`E%C0T;9q!q5*IS_%+<-74$glDa-2av|duA'
    '_SXU*n@VkyclO4SU+TK%10&k*!P^AugJs`}MF<-dpK2;Qbug>XNMF^%xLfy*=4#*XM!e5Q=1#ooVQ1jg{kq#==;i)(QN!}d!(;cz2PM9Ub9JD!Jr{jv'
    '2@^Zw1dA@_k{CPF+9VN<v7<ePMf^sN=o2RK4alu%EipSA@i<}<vcn}7v)0muTeCVhn_5szwYL!QhsH=(kZYKiQno6vt}&ne)Ht%+^gGqmO(-EdY-9ps'
    'V<dsV6Q<?k1UvG)Pm7zF^JUm2^SaD336<SSXnLB7`RP5riw2MZ1Jc_n<Pj1KzTk&?^HUAM>_^gM#kkbac+^!UZ+KC)RS3Q0$EULxQ*-5nzquhoJHqyY'
    'F5VSgr$fK<0z79wW2#hBg+4IB^mtm(DDyFAbw1qiS-A|wu0pZjBV`;VT85#0$W1xtl}+`GX=1loNaGihw&lNH-@!iW*K(QanQqpJg^Dd$iwRCr6Qe_7'
    'WMZR34}SKB_@m`iIu*OUF17ZIF`PDo*05=9k3)C&(35bN56L7B(&ky;n%DT$yxfY`MJjF<C}rIS-hJh<7N@&mYdJvrw-{$j!D+>-!x0u$-fIj#6hFJS'
    '>9VqMZIS)MNY}G_U5J1ETNDyVva~3V<duA?+74$lM%p!G1je{%Srfk}(dUqblaNBDLnYcWgmmyd(-fjx^vnJZ&tq>`10CZ)iv!xeSvBeghH~dvc~o1g'
    'vR1H0xd%7V@E!ORm$nDTs#I)Gq--zf*%AMHy#SD#C<hWIpsA%QeL_*yoX55F<WN}3B@)HzUWfuevF<kuoA>JYxR7+!2TOB{Ff4Cp4!C*7mVB-Xw(w!|'
    '7+`>*(9XLOZPb#xZfo~la5XyE&GGr<ZX{|DbPF-uQ&Oy@1aO4Oa9ihi)k0TQg1hdPM|*vb_hBilr1(?(58<}8rLzZ5^4V2ZwXKK{F$rfed{~alhf`0<'
    'v&An^I!j(AQbTpw5UW)u<5g<=R7Of;t<LQ4_dWBMbpxD8mZ$EHo&+fVl#(0>fE!1*4JsC(ev(7fvJuxI+9xOYhwv9q<inMp?Kd`eI}xOTOa33N9Jw6^'
    'w5p3=aS76wKJs{j()9@;oT`1(U<^z22>|wAV8Gro!+}k*u;K2e{K8g;q2dRyFm736Rw7u;DrE=UV&ha;(7YCM)*R2jhw#LQ3cx<3t(`H^(Cmv?;TXw3'
    'P|Eg=HBMJ%A4D$fWoQOhP0HW8v{1h)R`zy?nzAx;Ag+^WO{{GGD>Q)eoQ>}0H?@pb!o#rIqEjfIamgZpvRrIL>!Y+yCBzhh*Slf^2ak0>Y+ceu*;deb'
    '#3dTIYUa#oUyz#SM6rQ3ywSBSzK~uF*f2ZMtTC-=U~CPTVGI_p234w6c=wcZ#b3gTbHZvrHU&M|u9Ue^^%2gs6n>>!7k&j>3>+o^QP}(k8Lnw&F{?*o'
    'A$4_AibkF)i&^T%V<Mu_l7nb1=1fm5_ju85<Gbl)gQ+AAMf=$BSEPZRt9n=2uR#5c`--gaS{a|M-B$$fXZcr{)!$LV`vf5&zruW9Vt|fBIa`n0N=^Jx'
    'GVo=|3%Ib=JpCZc8)@7e)qcVuM|U$d#dt2uB*8ItmQXeP%nWeBUUSA1-uf2Cb;IVtwM^{-|LFT+3utDF?wQ1>ByZkhB!_)*!1<vqvpCI<vYoNV(aF@v'
    ';Fa1S#?bGg)uG>K7KrW0G~$P0Nxd)RBW(M&!!-=6dRM)vsO$AQ^n4QKDMd71XQmM!scB&KtQ$#U60fwot;Xn}({_Aah}M3cLla?szfKlB3A4lJ0&vRh'
    '<XhfHqSf2(VbSXCh@QYBQFyx_2(=nhx0ez5YTC+3*c}Vy^=QoLCJ82!6>9>MfU1~pm$|)4mbd0v)VGG0EC#X)q~JX_7!z*a4EemF&mg121Lvc0jhghH'
    'nwF=SGOwdeKq3RWkNq}Un>ok>G$ss?hKC$ZVLl7BD6vTh1c|#`Q+u{nU8z~HV|NdyNQ|kt8{Q9hCveN!a3#yzn{K@PZ#z5q-M-}5ghUOV)U7GF>6Tkr'
    ')|JXu*G5D$Vy0}gV}pCR;n+U=uF^bK$zuBmgUe2K(?-i6iet8A9lEESne&GZQMZ$25SA~AQXOs~RL$#rzxO<mTDLRX4HC!RpN<&t>b85kTk~Ry{$J_n'
    'aB;xTOd(!OT+@mgTa+7=$9Y|5)n)xOuCAO}G5y$6e!|PvR}-vl#1i`hvJL*g%+Pks?}scpL4=S2R$$@>q&v)w=?_DX8cre4>-&O8;qeY=-QRs5ONXM~'
    'KnQ$oFr>YzE2KzvH+RZz3)}uK=22LX`a)TL4U}XpXQFxaQd!AhX(MH=y2ahHEF;!TjYP)6WgL-=HCf_qJ{I<&UtlYIUmQ>Cb@Ak6c+rM8kIhEbjaS1H'
    'tT({+PC|vE#GR*!x1?kU>!HRRzS?-+LJW#g-aI^JK)=vqWlj|<%SZIMZXbocGsZyBTSMeTsqljF0`!<9<v$S#lKY3PxpaRpDuZ@MRb2v@i~Io_FTV=E'
    '>r3{rHkvcm;myaMt<M-ovcgtI8re17Rwnscn63BS4!eVwA$mTm=^#{Sd$9X@i;5gz?%3}*MAE64_@5Sz5A<H6-3c=U%^fZ5!>0c^ZyEYMq!)NbPI?-L'
    ')0tbIT6eR19C7mPyR+)Nm~nZJrV$OAL1i!<GQe7%aP%UPJb1w5?cWban{%G&MXK6MziZAX+noORuYSovbbtI0(t!F7!{)~VcgPMk0+e}H3U;*b<ur|x'
    'o}YjISp4|$pFjQmV@NlZrO%9~g9rzO*l(@-_as}pgu~E>7ZXoQc~i=S!Cn!#Pt7iwf6JlU$b*+q%bPMV?t_2YAmbQD7<yJ-e%mM1F-VBH>=7I15D9*x'
    '_4=5l+%lgh0osrr2YBlDyAcz~hUI{7e-7!%FT1zKW(?e7%Rh3#!P|%uaSigiVP?OJq;IW<ZolHwhHM#^yAsa$o_`c(1}uq9=`R6cP}Grm7~nw`z^#sa'
    '8~cWm48cpq4FZ=pT-WHn*kq%@^O52aFK$}flHjqxf7`m}-DdA>3=#)u*lrxpZP~Wunr*vbM{;H)q)i&l_(Ua*s>ID<mkJxU(#On8DhVuv!v|LmBoZGt'
    'G|Nw$j)D~Fy{sTR?|fiNxl2TmsP0I`K(1Ph<b#I8Th|`zuo(-*xgYkPg#Jp%lby_ppJUi|+Z`U>Bgs9@Y6nyV+s&hWX`>8cT3xoM{)0a?L0KaT#iqq@'
    'X9fG0E2?Z-Yedq6^vaQ3+Y6#pusSc)T_x2G1suTbA67J|M!ZyimQBpnsOtv2dh2y)r@k@%_2bV!{r2O>V*lk|K4!|0I&j}GJvt7idnA*`DP}5Yw436`'
    'FF$|z_0w;~Z(o1?9hRNnijO}RAOHT5Ka9uh8V<L$_QcF3s~-W>0p*QQkhRAvPRgv`q1KKnJ(n<UA5k4Rl2isgN6vW}_x)RrT!01AwG`2B!)UT;0+wT('
    '{u-7D+6?fVL{IpzSKE+sj3j>)K6Z+3NOs-wofC8Osg(cp52#Q<)CtG_pZ+0Q;$IQUYv}1y02MK)x)zjn(;P2wr5Nq0gBlv795HDVL}<xkZE#sViu$oh'
    '-jk6gUHbQL+6?TuclmJI^+$Z&pdaIQ8HHW@w{sP2#S6Ch<S2_%{$yum-Hi#Mel0f*a+Vj1J&t00@)2hF**=Op=BeVJi9Tu))omf`5NAm~;RI5a=OEMA'
    'WsBEV0k59LJ72^&a-88R5yBH`g`2>sv&HjK*}LWHu6M6w;QrEvYk_<(;Elp(_cCWtCc1b9eE=TW$AA0v%QGKicB{s|0%Gn3M~zvah|Y&1_X)SR?5pk$'
    '4xc2WcqyU^s%ocB3#*8h5h5Md7?K=BaAUTxFYue3W3NH}7MS`4DOtZ^G{==gCt*=J`bng06&@c$f!1_?vg?u5=(#QfQ(-Ic#*YJsC_k;NZR^v%(qa0@'
    'o+{eyIg;T}puhxC8^P~;@CiE0<4p;yxLG>VKaIC&Q?wNxnnu*8fT0RU!{U{0pw`s<OpLO{=XcpM@fe(98g1ZERC;2)2F@W2yhY|Xj0S@?!$+pW!?4-F'
    'YViZ6WJW2pHrGOxa4xi03nwcauv-W`F8NeB&B%ZJ7rp9u*V&GkB`iS4At4b^ZKFNMzNpLC+|vlSFr_6d?aT~>XKKrLw1Y#w3xvbmtt!F5$S)eg;nV(r'
    'r&}*WV6T6?-}A<~E12oZ+q`oL>5Ue;ucz_moFJo&_25MqR4%Yb)>UPP#t!?PU&2{&uCqhS(G9*@oYv^LhN}+-gO-wEgvtpR%mW-G{hqak?Qo4K5s!2%'
    'sl9NO{mE}qVBVvdsh<rda=NLwBsRoD6x+G|@Bhhu(GH&^_QYeRr3<lOH2ycFHiuRZUD4okCC{AQPPf((K#@~gt=VihV6vAzmfMI3HrYFg-n!&dDjV$k'
    '70!;?&|luKOT)Yayk86BOi2d!VZ?J8dXX@wU`tsST0G18*-PLuhH~jKj8VrJ>;4fAWiS%fg;CF{de$QK7aDs3&9*wW-E>g;r!dhvE&}TLJ^L6jhLN@5'
    'ry7Q7Vi@SBebJqc;Kj0tle{(SLbcCIJbR%>55p6MV6>XRM3eq9VoUst_&`wla0Gsa#3lY7CoNtNpGX=Sw(H{E>UMY9vX1!u301U;s{HT&SuAfCw~tPM'
    's5L|<C#y)!+zzXKx59S;-a(e*k_n)f8wx)<k*0Yg?FG_m?3Sh9vH6)DHR~RFd*{T1>V$VpJ{Ut<x8p!J(HLWNfDI3$L(p%2d%`Hb3I81ue8FJ$tg525'
    '?!kd0R{`c=ce<HKM~u;cVwYvesKhgSYeS9bw7}eYbYD#5Hnbdiu!+tk+Z*Lq3cI2v=9DwH?CkOomK+&nlrvD}a<-u$?G(v{DcGo+j4=32;VNAQOmQ71'
    '1@n5qLylPKfT*rR<e2LDBHHjLq)Rr^Oz`m58U}XpFhga0Ya4T>EgjD{g31>`Gv46~xOJ=jC8@4n1a@)l9yuR7p_pa+A{w8UUn--^^P{>~SAZV6YZ_G4'
    'q2l$6sE#K1T*dvY!AUo*IW0W2#lSi_88>i@M%CiR&D%=K#zQR0YJ2D6TT!-PM9DWwCARy%!I!5T&ru&Bo_PT;Q)xg%HAM{Uc2d?aAg+fE#<O|2Y6>S*'
    'fRiu4(J^uv;R!RrlV!kDFX7n;ZI7p?<twXzCtkqY?P2fR#(-s1Ks)9<V}TKE7$oi%DPFDIz?qH}<mDH_XU(^K?)FHVZaWlr-Nya`dy(3g$YtGxrMZr('
    'w&`|<F0B_U1|!M!_748>v*yS`)83TrHb?dB^Zxj85!DW2ZDvYXd-^F=ThHerRv5scghJ*etlKJ})-R*lHk>P=&GM(aHKvYy0c(d&G3>rc3F~dXf_e!G'
    'B^wo?e(3N?+Kr2huAGKN8&%JgDlC<7%w6J;#rN`ov}GURT(aD|`I*CLe_EEcIEP{ZqFrojFZxWP%0h@n2gK?sM0;&Ck7fc|eHGenz~{wTtU6$|S7Ffx'
    'gYQ!Nc|=DTK^h|Z1<1pA5564i8Iw*ZJvF{5qt$U9h6)l8N)&Ju%I}+%3geRp<8mC|4R{hBcXX6II6}*qp@U&$6L|Vsnc=`$cC#Qwqbrbf>DXknYU>Ho'
    '=}5wONKJ{kO|ysvQKn?DZmEDDO<q8vAvryD=g^wGwq*c|21GW6&wvS;CFe-{MpQeh7eu3F3Ll_{cW}tvpZdFVVet;dN{rR+*7yR-1Ohg;S~#vcIEz+N'
    '$rxgd8TCbzr2N3aIxR*eO{0~K?LtOqts)G}8C<W6yK$9fXJCp0{7YCK{|1hea#LluP;j*J0RfHQfEX*pTrSp>Acn`Lc&+0a)sE=3ZTJp;_wl#ifBG3y'
    '-0r^o`qMAJp;Y0=zkm7V$De-x@>{X_>AAT7_3Q5+e=e3E|K(#LD(JbtIZSO@d@phve%rdcC|4+}kw8ssi}kEY&kFCm>zp^-sX^CT8=iXh1<+mty&`gG'
    'Z1Ed@#d;0hZ=M_ZENrml!yODk+(wR1aJnA%y}f6BvRB-KukQ)GXI^gRWk6X|JfuCrh&^U<=gM=nSP^72#p@o@Pv6kc#zZ#%bVaI;jd6rPn3KY0E8?LX'
    'F+yQCK0;Gk<>%e<P`jjOyiemnkD;eW6yL&BP#YV`e@<$^o-~XmGXFC^>QgKZg{-I7G{oVIM=hPvBh@aotb(7FL!)jBfc8T1$!2ywSFolZ!_^`~LFuJu'
    'G@^woKt*$fb7UC+#(0(9uFEzg|JU@ejcvVzbHtjGM)s%0kW*4LAr`0BI+*;^KiuI`2HXEX{X=e2)xpHXCOhPg=nqipU}PWEgn#PLhVKd3=(`RU^mV4c'
    'pp@|G&q~{|!+(yO{_5022L&5PeIR_f8WEL?<>=Ydby>x!sWLvfR&cV3HRX1jzGJAetb&2c4e*{CpL_GK=+HgV<VK;Jb2AfttIEx?SQ$4GeBQ5Mp1jls'
    'M`CVKMNQD;78oxCp5WcYDllZ!6b+VGdG58ypb3>>KOr?ZOuD8DUr>{-nR&8OrVc||P5Nd+)5iU&Jx@s8L~TqLL{H#39iP_ucSPV8;S2|76K+Z8`V`6O'
    't(-bdce`e@Ex!SDHavpaTkec#{9aKRyU^^{o?+||v6qG0?{#;A30(NcQ&`=1{njP(OC?xQ$4M0KkSmt-Jcu^4vNwJFZ=^fHffbM3sTBuE!e17zXeT?4'
    'uj7y>5kj8$6?lTM=i6K<``FHr)4QLk-S&Ps;vQrJe3#&=q5~!FV%EdTQ)%<hix6Hyuv=gxVpphzeG4d!^Kq3WD~f?a)D|`$QVc|mu0o^pjrdu^O2vU<'
    '+GgYXpVkoeXkX=V+y&a42E6YY)sStg)v8tQgKriiQx8t-Xo~{-T(!ucN=Jx-+#%%zYJWPo!2AehvA-8_=8AOTrl@w^2;T;WNgnmHRK-7HlEAv7yuQVD'
    'QIIt^rzuBrM-Z3Z@|ceXA7D5J*u}z*c7fa7FHN}%HtK>6QSGMARnr6>A3<KNH71qjb@APVVe~t9AL_fY#au+^J@U^4gp*te2Se6fI+35Z%iV)7KgA(B'
    '>;35iWg2g~;sMr}6~1OTS7*5k4aJVm)X2Bdau-_2&XCC;PaB7{P238|QrynY=6Ebeev7w$0fqp{a$Z6U3K<r9)>1eFb{O{k7T3s2V3vuS!Nhqm@F?K~'
    'fQ(&-ANu|{gSor!tqUw2S%Srar5fsll-0vxKi<RnV;$3A+Wpl3v{s3@7`o0u5G~s+VCZh{{FfVb27^NU`Q6)_h3<^U?9?m~i9*Aus%3l0GF57F_M!|!'
    'mAADS`yEm|!mOt^F&el9ng)z8a*enK1IBTNuwR84-(CaZrLO`I{BIMtq%to78CM^kh{gR<Mo6JnQeK0EwQ>kpD6?C8g!(E3ALN}DWFOk*DwxnVQE~QR'
    'Wa{b+id$E0PzD%z>o82Dc)>bt(Ezax1!!J?;FDKu9NIvYFBtO;T(IqUNQ6`E)RwmuuVBX_65f8(IRK76J`1p5^CX-e02bb!*-&4!cgGjBU5CQASJ+0h'
    '&S*V)>sno7J=&8N*e?4}>iP=XcozlpyYT6%y23nN5Lp0_t>qQg5j8bfZ=(h;udt4fWN030ZklV%{pk#hccH^7bA@rfqVI>$W!KbvKfLe5GzpDvuIX=o'
    '-a}h&DzEykJNcpMGSpdLX-n5UYd>7<20#$)wGdwV5CD$Rh`FZIjz}}rfADp9zX|Pr&6wdw{tZgKZv#j)=+)IA5iKHOL<8WO)gnG0<Ft)dyNlz`MFB*M'
    '9pzOKuTCjl_JS0bXTQBPdlbd}4(&MRkYcHk6v0w&&~0S0dyRaZ4`-=6WKzO3v~hjJi^^wFX8KN|@`rDJnsgu(D#xjwLr>>OlbK2{fv%bC@~r8iTRatq'
    '2a0Ll_S^m$vk-Utad&vLdl-pGZ9G1_KHR!svq*o+tAN8t+DjjqsqlF2L*AQ}VegGbCWlL(^J#Yvi)G~6MTL6kcDtuyG3<{Icij{8<^C|dPYxu-!$NzI'
    '$TN%A(laQ*^@uQn=OOE}2%-GW{dibNhY)!m)inrdWjsvqKxA_b47(u_KyWVWs{p)17hwcXMz+^rgoXAY>Wg?gzijKe!x!K#?ox3UCUK`MhO*}+KB+X}'
    '&=R%AW(c}#wkJ1B8_~O2MhBgG-p03CWEGvNs>EuCVoCdM=8G<nG@1pCxojJcng}3r<D2#}0&D*`Tdz+d{dF9whD7z|r2*%>T_65YG=<YW5$Bd4skuMk'
    'qp@5t$G8|zT4KZh&wMFg?l3RNevsF=>PDUWOxaYW=K*4hqH~|249#gh4zE|Yo7-fiM@83K370?J*0K8(a5u<6ycodWw}qvw%{%mhmgCWrf6H1)Uq<`W'
    'VU6qIzT2WK_OKqab5y@Sj9VmYA|14FJs?V}g@j@+u~&E<n!f`Hx4Q>jJ5_91)opxoz7q2|EFX&gIDj8>#5@8U&%9ZLldFohTsNU%;4Zq;aX4(ETQ{(Z'
    'PR{DGRk&KRO998agrC;-2<s3DGVC~6JdC8BFwuE5bTn**=@nmmEPxYru+isjrZ-<o%8*X1qWdMh5Lv!EcEg%^+xa>78DHCzOC`+kS2zj;GrX`Ehx7%H'
    '5|f=yi)R_aAgvP<8h}?WZKz>Y+T_#lB&ow?75PRk)}7**i&p}<7hW@K)7M}AdYehCtLWXOtRf%|Ywqy6cDpLRLkh0%B2jyWyNyd_LaExe5-I{a!ct*H'
    'Q_|A4^M$M6{HDOAcS#hrOH@PR3`{*)ccgz9@<VkQ-Gsz3cSt%bTqcbW&N(JR&JNB^lycmy7hUm+xJvQoYnNaC7J*IiT$Mlm2+nGb>tE!BaNhqlxa_~{'
    'Ql9}9l@7q}MDPMIy%f1oT<E$vgDsn;4A}Km*hiR517^8&YYxg<j;HN$|5RwL&Q*ccOzmf?KmD0CYFO*`Yzz4ple%obp%&D$2N*c;TQ_IFV%+5}R-*id'
    'ulNHi)TKV#+is7*ugE|Bl_Mwf6C+2D{6R!E3-zj?bFvZZrD5???VtXXJ$N2RZa5=ns)X4CjZf{;mfLFbo&GQQC^8*dq-wP3_Aosv3?$6ZQOUa2{@~QI'
    '0*XgRX2)@D$##e_%kuT+Too3{tCt~Cdlgc2RTh=VXArct5m(_%9`w0DF6^uZk4awzH0~$0@~2lQjk*dXeZuEDdGE8WO?4GYe5r7U%1bH`v|Mi8kao`6'
    '2OI|zZR7VwTe&`h6XiF+-Jc$ZE*IB!GG-{Q&20sPo5&&d7+o|gvOT9P8^(OCOsl6zi}k60IwPlTTA#9HYnGZK8GEkl$r9L2YrF)ububfClER)VFSsFI'
    'n^DP`nqwZOC9IfMmOFr+i>qeFc-1WyPyYwB+PGd%bR7jzQ@J2n@UiTbWrC0@9H;uE+;zr9e!>n=n<+w~Z9K1LTm|Y!JyQl&HO`A~z`bc37fIX5wwY;n'
    '+Q6)=v_EDfN_7n5DOFIAVh$##Q@g9TOK>g%0pZkTI9|-V)v23B<Rcr0nwh$u0dtHGhhhcp7_-S7yUh;$MXR0Z*-qWSsap{??zwJ<lQGAli#ZWn&%AFX'
    'FmvNMF<bhsmQLcZGD^>w!g73?4t6Vq>)bIE!krmtu^{gB)5dCMlL%@vGrNNrYWj7nQtG`|j2*70=KFE69v@wcS9yz@rmlQ4bU9OcSkn3>Zo!eR<m_y*'
    'xydYhItN7Zr*q3i?bEH0xq8m@I2_?rI#++u`q1}=={hpo@V0-q@4L;xwuEj}qE^`kwaPXE;XHj7jj%zY3kp)nCVV(`!@60=6Q1UjPccM#pX62{AQ(T4'
    '0Kw!4JRK+I6%0a&!I7ok{i>#nYq0Mzf1f%KlYZc&l}KXR8y?}=p{=T>H9ugO<$h1nEbnf1<Bpq;tl#O@IQpm(WzEQJKN9<ejj-7M_x}=Q@oIB}bRVy='
    'Dr50(-`#VR@NQ78i&4U)xcV$Yl;@j1((nZ~Oj_B7_1?9>1gTCm#_4&FPnQP*xVaBEJ<`jox}P(64cBP;Y~vu>0MTBCh#y=ASqkqlR`zW?u{_vzNuFkx'
    'n5gLj<>1_)&)bE7_vN^ORT`s^Fvk*BX;jJf7fM)<Qn{#5|MLtg)dbcmAzY3rb{+M-y%3U>Qo$E)^b1(Kb<2k_ixMun25_o<K3ZK~XrlEA7H<qY!oZs7'
    '?oGO_#{HbLZ{h*7+emwWe$LBd)!n5}1J^PhGB!#t!rD?NWS2V6g|sA<H8^gToMa7u+U@MUqpm_Q3qw;oAzBt_fZ5TO+0a9H1RIs}GC~3%%BT=hUjacE'
    'K?%qvu85HrG4dL0;u?sTOcTqb==R8Y#&%lf4@sHTtT7FDIk5|7`jS~<l{Ry8&Tdl9#X2eHY$a_PLRM0=42IQZ!^AFa*Y>>Sq_oREwVnN{Pb!~UCv~ka'
    '8q0}kuxwtij11G}ovkam_1X90wynxbR<Th!2+o>K8YY=v)$IK8X^y#JuoIW<VzaatoAZW~qHa^u*|OSfnTcsqOLf6g=61nMUo?xX_Gv+9?P!;plyjey'
    ')0{VyxmmEPE?LD!X`?jrW`~YRPV+A&EmOZ}o0z6!Rb4cmG|?fkE@#{5(5)%4e_|WgxGpc5%G@y2w!Lf@nZ56l56;e#=-ixr6s?1?eAzg$O~+t;*?Q7a'
    '?~`?*R_D#F7dt6@ygt<CD*MDd9VFsKfN2#{gbVFAHBBmCU9gn7T`<!Z&0?z|?W_65w&WhuvJOdM+l$6>Vj3*#7c9M@$vwNNrJ4GMNx4nxpelJ4g3JJd'
    'y}Hsq)M1e>BYFnFxiD$T7fHb{+R4-`SeeUKp_I$CrORr@Fb$#02M9nU1-=SF<^aLIzRo_Fr=CeU*Hi4zWaUo0C)J!clbKa8YA+cDn`}p!8;kQBib?hp'
    'U0yPjnO!gwm(7Azwx-Luu8cQ?tgz{3q~%N2GB*yk>SbGNy1^4l{9O^Z5)<KtgH7~JG8EA_^#wZ-TaC%*#H`u4A17ZM<FA?XX358r$yGM9R_+QQ{oF48'
    'Ty@DX`C#H9Cf8QM&ia<=_xADk>I-)1=T^!0&e^%(yvavb@mHIdtiz|rt$vhYsL#|>+fG!a`<x!r`C-l*Y2&}NPrjteIV0Pk(Wg=uzr17>eGBYX@yEox'
    'S!m7pi>vgLFWQFhA|r^fwmR5ao)BisW4U@-GaLnm{eySGuDoFPuf<K#tp|J_v#+9t?%0KB9Ah!^Cp7lRcQ~Rp4Aa-*4Ku1`VARtqyjLWe*IM0JxeTMq'
    '8i3ly$rvE1a23J$BwU*lTb2Lk@4pp7e*T>4ImpHnbUU|&5jNdwd)gGQ*62*qD(1=9kg{2eI!Hx}R>yTXoOb(eF>ddy>X9#npi(rp5WD3gzwxn@FL%Mw'
    'MrxtFK}HjywZGrj?ztNy4~KrU*kiQGBg(>7afZ0CZ_b<`F1!u$JTAsHu}flf$9VKC5WA%krGL?>IByl!9SYre=&)<vSV78t$JE2OOJ=Ct_AtDqZcLK0'
    'tE|IJVsoq(Vc`-jQF>cPIV-2@h5CIM_2pR5b!94p*ID5`ou5h0fk|x@zuj#trU6@V&m%KFyXQXlld*{5i`h5EHPT*erHrj~pp51%o~-U}``h(xv3ujG'
    'qkVcbUL%{fNbbb(lWq8VJH_D`8G3hF4*TV4&Cz!;6=sy<E&ThD`<cQ~7sb@#%r<@;-WA939Tl+>ap{EgrQ@5G2rglq9{2r7_DxbE7^96RcD7;F1aUvC'
    'PHw#5U_?_@gL@v{J%oQvbw6iK`GUT#+QJcfYayM)D1Ow;TDs8Ht*A^eZOY11MpgW-Fk|_)+usfJ0!A^-ab<P*?!+3pa*IM3_!ie=*ul;uwXH&JtK9bC'
    'gnB2K8uQR`cFJ;OfitXqYivS;G;J79t;|ul%Wk8#J3#Ld3|SJ!2d#owuPUSYWl`{4{V?taWN4_YG9uMqwpMt>|71HXJhb<J+!w5ii&I>tQt3z5C@#I!'
    '1pho|-xPYrUbf2fuk+6{_6A=_rq*sQx}RsPYuH5RtearnoVT|Bt6>G|SasRz*m^2FjBgB;Kpu1qz`VD*QIUkLY@QXeq3@O(G`Cb$Z7F3K%x$PptZb)Q'
    '73yHc?Pzf!#j~B1BCeedwbRWk6s<{|Rul*bXw}mynj@av&HbDrg^3b${G1>Q04|VbpN}A5xe(DJrQ|FikNaXGyvAS>KS+Ah`Qq$(M(odX$hSVeJ1M4>'
    'zjiCV2rIp~1UA0!h(9{T646>wH$K9We@?L89LCf8emm@=#)O{kEKgY+^j?ViQPe+vzFlb{qmI(XF9Q5C`iS2|RJsmJ)s}7*hE>Eb{j5np2N2Ac9qiQI'
    'mMbjk+PNY^dGb9NG3_NBc0c|41beolK-n!@ABKsuL0*QzaM$gK_GK8mRcl-<w);5`;UC3WNtN2ga{pNt8OH<Ze$L^9j=Mv}y6(gz2};<tey60PQ-S_z'
    '&NVP;Z&Cx_q2N1g)(9f&G<XIU8$ZDtb*@O(Ur3|op5@ouX*7;x0!|Ak&H(bXK{eCmJB(Y5q3G|S1M8s8ajwbj{3f4~%ZOW(s=c_Koq{fS(*3-E6#{+='
    'WDuOVZ6-`+a}nt_3Spy>Fa;~ao?T|+`=ML*dtO>upE1(SbdJ{U0ovOwg6Otnt153nVK^Q8M~rB_qeIECrvtYv5F#R)(Rjr1o3T>YLL#diJ0&t55ExGH'
    'oeC-E@X~RGiVI~nBp#Mp*6zq<U|GqUYX^=p{69c?+pRf#Yalm`q#aQvv2{^Pc3*SSFYadq=9f?(p1WcWajKm?l(_9NxDL9#2oqdZsS^Yp5cC;@<q6zB'
    '626ExR2`N_t<G4!J@qRq-1uhM-E%;cY8vY+Y^^3L_jbM{c@*|t;*UFCM*nlh?7&e$qyxMMj|a8RaEu6~#J6ED(%k|_KN-p^fJV!gJN+cqRj{t6)<=|j'
    'VRFWK^D>y1=gb`&V~)TJxe9A~D1LHEVmsL74n@YbMTdWk4tFuy!czKMA(178!s0m+inQyFNO{@q$L*nbrOIbn7O&m7Zpx-HaC$eTmtg_FQi}*I)uCR}'
    'yE^L1`RNRY<o(VF9Z6j&B`YfjO1B;M#)dM>D3+F!IyT*ESULwQy}z)b@NTqXW$ieJ^WS^~d=qB=Vm)p{%vaSKZ;Jm0WmRvxf+ji2ZrJdyMQN1@$;P>G'
    'nT@idOOmK`Yj81wm8g~0=d8Sdgo!x%ZJ54y(3DAvO40evOz^4=z@_MXw+S3a8<gM71m#!42_RIpsiSY4K74=!=IY`2)OUN5f6BvMF-_5ZO*hVG##^Uu'
    'Z<Fale5Q|)RmN?_((<Z$Bg$d$K~qTEzJV6f)s4;}j0g}_4&ei4B7`anRQ4J=l~vnH5$t)PICS?=oG3{c8?mj+=0@a2^4$ASB%K3>=uTfEmE{{t=jPUq'
    'kv;}^fJr<4&?7q_>`3q7j0FV&sH$mlq)NHFy|yadg-1Z5fY*YpsUR|I-MV&{QWffj=Jx)G!c7A}UvDv(yDpooNh;?|K$EmG$2sB#YajJkvcfk`J#Ael'
    'ywDo&as?>+?i~rJ%chpfyV+8+jJP#IciO#(y2T13MM?g<Ec)BdD?TA@bm@n<yP@+S`P*bj_U3ZK;PkFFM~GT$mrcLyDDRml*LK6oHv)0fhT6apsMc<a'
    'L{C^up|A~+TNn{w9cZG7XES%K>^tgp)zO;i2Rqlv34aJ^(99zz7%hc^X)3pRPhmE26z)EB=RiagftYh3HU{H#Sd9}Px^&mlN;2W4yKX}#-*JjL3tfr&'
    '`|44a-=hxscW(ZChZo%QTzzoRtiAMa6-~wr72($7`@ZP6-G=YcE})jEdFR-tbt5mJE=H`kC3?67*wz71y$I-eCm1pVYuYA&l^4L=sJF4q9`-PgTN}^0'
    't?T*{E-MxAT#DJ)=TNPLDWO*CBB~{Vg#*EI9QMQELW|l3lVNpv6|5CvpDm^E64IqM%~f!RZi_cw7kXVbp^TNd3iLQEP82SQs9BZWjS`jdx|YQGvpqA6'
    '>M~S0@+YQnQ=%mfcf`+F<;mF*C3WMOG{FtL)Ja5NduLPcySwv^s!E|;D@cD#oyTOR|I1TTVJc9-(=Xv=GjkSEiYg$Qmk{H}+B1-_lhnS_mAJGZ1*(u}'
    '@sZJbacLO3b>hxGHOvj0%DR#~@h-L3WMP<Zeos5^QXbF1ZBL87c%rqTxV$)^7+&o`+e_naciQcqE<nK$uCyP*vb_Mc@7Lp7J6CX)Tv`|c3qDM}i=a3-'
    'uId)un$Rtm7Vn4x_w3@Ty{upvdqY#ODFyCy6_7$zVep<olgjQ_$Q&Bt_h-<`+5|Ln1zOAtZ!zk+2^ePv-9!J_t%yP(*-L$N+u6BlS{=aatH3X;0lUL+'
    '97eA`Dw6q2|BjFe^-=1tkr6DW-cB`ByTEhe5sGIA5uH@s{ab75=%fnn-)dx;r{crx8Q1m?<wuHWOj^co_8dh+Qo9qM5xsIuO7vNp(OP49{_vcK5f<o3'
    't19uMIX&0uSz~%`6}}%DpNrsl>kemp;J^#gHu08|USH0i+U<P;J>8gSdwF}>cZb8UEPfdG*zL{CXFZi45@3XWe+Lf2W++fmbv@plxn3|XJ-D{QiEK;v'
    'Np}9jAp{3<u&&xl2)>!}f?>vn>RIbUzt7m4DO~H9zZNHH%4Udlg%qAojc_ULHNA;=sA1cKN4D=b!wFe}pE|!+%L)fXRlxD3ZSHb7pXn`CsV>5C++2dT'
    '%k$7(vSyPE?Tb0+4d<Xoa!lOfjk}(~9Q4pzKJ|RImxKs|BMeaYIx;gTBnN!b*|4dCC$JeYdS0SIe7-K*rpC~0FGXsquw;(Q;~39_ckg}pG0Z3wQCy@z'
    'q7`29Exyi~0kR1(7H6B`G_ndf@)C}}a6p(?5tS-tgqPkk(QQM(u_0l3&7_>vtWu#w?M0-xYEn+B7DQT2bkdsx!n!$_N!4W3)YV)QJ<#P2z$6!O&`hS='
    'WLP3-C>qmV!g6}p4k`y~LwiRDni(LBdUPOrySh;NCoC<>Ybm#F8;WChR+@5X3cdv#1z?3aNp+)Cc4MPkr_`r}X~XU|(PVngC#n`QOu8_LBf^B)sp`T#'
    ';HW6VP;=@8bdh~G9JXiXt>SA}x}Iuv7#Hr6u3O-%__<pcD=fqsutra;VFaxBooEVRw;>ndgZn`kPQ=)p71;7bD(x8(Oyp-dJ_fq|YVuXEVeJI1qS0~K'
    'ftT?F`_<uuyynA-HrjC0&r$AfL<k!+cG|=NR2)-j<Kgk4P|*!wx7&}8x5dX_f4{|ae4>AIQ3N*KiEc-08;*dWcK*$15Qw0y*Sc4WTGi}2jrf*>+~B0L'
    'yy+eXM=QC9^WalKBnPq~g*~U`E8=$ontVE3-SEX4xP&B2{y-AG-4ER`c-n{D*1OTARTG)9oe>SP)>;K?bH+NSkFutX2&Uy(Qz8h<O~XOQ25!!nqPTKy'
    '=at29MLlbGnApiS*r_visqT|C^m<J8L^^Aj%0IEcU)5C`R9<9mx`Q!mpKq(y7<{~e<=FFPl#Pgs={y)x@!lT$x2SJsPk!ASd-*``btj~>zEN$sd~UTc'
    'RqMDCwSCnc*Y8J)5yGeB_)Jc7xIhcX#w%T<7Dl@?-WcC;;f#lmOWpbdtf;IHpQu8Ems<z1AEEgAkj!B{U-lGyeC(D>ik9)CL=Xp~s<r3LsI>N$;N1@B'
    'aB`T_@L>sa@PtHMNkHs#;S{v-`OkI^N(h#45c3k}3O=PgGRF=ilEz%MXYoe<>)&7kFE>l4Kt+V%Iw%?P)9Se8==Uebb}qD+o6_JMar+^Mo@L`?{@Iv4'
    'Mg;a4j?$(-KlFDQv)t~m2mhq;4t>Q-<VbndGk@NxwyHdvtJY2S*NY+5UVlpHKn`u?BEgZ?rp%0EX$6d7gtfJW{fSGbgD3|2ZSPS@@@`zLd-nh&n&<Y}'
    ';Mz?93LnBXhCQ~e%MB1E?@s{05iqi0ok{z5T*~p-t9Qg54QsA@!XJ<RL#lW;VC*y;U_O7gjD@?!5{9Xdk*sadCt2Z1b~tpaeP?~o^-h<H{Sa%sLSs*)'
    'XJ40dW4pVcweZhO_<u?NRZ-3vQds7ezbk2mrM{Q3++Zl5Zo}C|qV3w*0<&8RuGg#T);ud_K$TySK3Qz-u;VL)&{d%hAG&)CoVn{)Crs=ZY?_<D3xMy#'
    'UrQf|4+D%Z^I^g2KP=ajEitUqc^XYpqiAO&h#b?u@Zy>7CCv975)qU9dw^_HNOYTgb8`e?4(U{dDShQ(*d0*ymGKxL;+`~=tRWoba@nR8U-x$yt8JI5'
    '{rCg2g=I?=td6m4GDDI_s5N`{XW&kE12Hg#Y8ROfHe<QNCl;?dkLS}OUWOFGebRh3orgSjyMfnH)MI@+(VBK^e2{|U<vxP5>%7YXN!!k<)WZU>Tl)1q'
    'z=rz(%i1}CwTnG9<6lnw0i!*4^l-WB68OIaZ=$xX+W70IG6g^1nyztDy>cN8PPo=ng9>&9cc(SQ=7L|j?>^xC$4d*Cg1dE3^2V+uF>gtfcra`$uji<^'
    'uTSd(Iu9KTmJv<?>@g!T-{}KheLTG_wGB<iLl|uDm9p`JiZrQXkR=8_MWk(ZL8HfNu)FD^(_9&V@MKbQo=*alJOWq@3!uRfg8v_`x$vg3jS*}WUcr`N'
    'jC!56gK4x|5T!Knm)(}mA_5Kt`o%UzEkp`cvA;sQkB#NGQ1^VtMR9a%%OH>B^oUaF9Hr6DtE)J%b{-{e6NkkoZDd)MFfd@J^U(~A)fuo@6maEB0*9W>'
    'MksM3Y-At&YaiFf*Nv&2k0FKsTZRxgex^$Q^x5q}u-zCWyTZV17kc0F)vPlif)6-eQQ$J>JfT41hXro#-5Uj7M4BE}K?C`Xq@~-bJqEiWn>Pl)jyC>n'
    'yTaTu-u}rbMKRiqVx)j<VwCNv|4_UsI!a|N=ucBmxp6G)NT5iC!AM2$!ePxpcbWCGo!`7TZHY+`g%Aq$sOo}8qrvD~I9o%*__$I-gK6I#hzmcsf4Gs2'
    'M;mVU3ZG5-<LgvET{nDwQOd<c?)w!*o_N{vAY=asW9O9F*IW#?*Du*OtJK^yOrJ{T^voKYqct=xIhqGjOsBUfQ$~xu!{K*d#*r6sVgOiFL{%ayVe7L8'
    'E0%ZI_it%+%_Tpjc_jbvQ53_|4H~AdfL6-I*=pW`Fgve@h1NO(s_4Nj#mL1L^K)`*=~5arM_4&)OyA)Dc9w^9sTZZoNvmOY*O4jjn&{ddibJ<R3Sndk'
    'Jy3?mV#Hw+HUJ&BuBik%yRGdd<Cqxs{q}C4xdR7?E!yv4^Mq<a<CMXeJY^<?Wn((zT{7lbyWc}Avr;;nwHM6zPrd6Zg@H3}9SUqrY;0}HM^zS{?=RHU'
    'J|2)Ejfl*U!;X?qmUf?IhX@xpVR!`Y`ZcF6-rr-NTYl$QKpxiq(%qB6Dadd=cQ9Kp&a$J!91P?=<rLqt_#?%yqBiuSFwvDIiOuKOjSSe$IqYTzTAqb&'
    'V9u9#y;wW{PiT1#8eI}N-D2teP4BDaT4;SE-_R{<B^IyK2&=j_?g~dYB}2c5L$p&~Sa$b+{9lEjMVhH5e!0yqNJ<|upoB63!A`bHGV`ITT>_G{PuTm1'
    '5o$bOx&c&TaXOH0{nv0=he-i-sa@<IIE8di3WEcs3u!+Z@Q!wLKXI*i)}McOH*~6eT`X>Q>GCCOe?3!kv?v2EU<Df9!O(GuTF{bTf5H5G`8VIy64RBX'
    'BPp*P7at?QDF%Hil<n=Al)g}3pDP^u7+2$0kVf+87r_IcfZX?Ma=h)s#tw5QR8kCR8e^On!6PI=IGhfL{?WgTsISqbGx$N4&<(8>=REJf{QUEm;;)~I'
    '?#EBRfBgMR@oyi0{aEZj{o9w{e*g6AAOFwC!q6~K#ph4I|N14IC`u^5v=fnD4-{36OT#bE6zadk*ZTW46pA9dVQ={SdwedT-j@D>rEWs!9G1{K-~u1V'
    'kP+@<M)7R7DPG+V>rK(o#myb9S74`_<tE#OjWW{rKi$mmpFW*7o2Mf5e&~3KF$GeTkyeRd>el_Z*+DU&eaR*54?Dbrc0Ot138E_}%}uE?F%|cUj#ra3'
    '8V}@Ig9Gt?Ot&kc-7!#GO4mcTeCUt<9L+uL@8TZudStTch39URZE7bdhkk>Z8MrW=MM4#EOlffMe63;^^U$p^panW;z?8efUZ>M_C#v)yH5K0FDOxad'
    'G#%vwh#558NBKl+VG6?G+kwM-p?x-=B**ap4(1*{J^bZ_*$cdzd_d=~j~q>jfj`Dg1;dQRhaN{nPtU@gT?)Z>#Ee?-bX@8PIC?}8Y;-=gTIaYLzNcn9'
    'at$JTShP-uC0-*(-Xl>Oz)$39j7N$=!9+V!jQam8`<5m-ZY0gW61FE&qap*3Km_&{nUz(QRK7Z2s-oE(vZ<PGj+IDGQIv7}U##sf*z2BV%=Wa$)*o9q'
    '0Cz_qA~H*6ZLO3l`~w1khr7f5^Dsc}6j)51I_<VwI6l@_g%v+m-?jT6<TJ=(^YQ8QwP@Sd#i!53$4?)gzJ2`V6zdF?K~=7mp<n_X0{WfKzlruCpQJun'
    '+0E|jcF^DKVJH(WxfsLpj&Lz11GulU03ix^D<;eSrC1DXh+!EblN`z~8OAXJKL#ioRTpQnp1`7ALmJqiG^&VjccsE=Nu;;!d<8@suYP+zJ)NE&pFdxP'
    'U#HieOzq9kYvIM4Q-XciZ~_<#FUYbNUkwQ*bhT6o)fMCT9v}VDhLTBJEH0HJQ2<jc7cz*%k*a+RRW<^rq6&u05XKLX==#H%TI@<#iW*nS&+b8-cZb#O'
    'qCJNM*~t`tZt{7?NWCa=S$&t`R2wZAUuz2ZbOUBapwqy`aVT`|_>x|B1S*R<&kKYRT7$<CpsDgKS`yeghUboLj=c@h_Ct8*qRvOVT4N}VfAO&~>;#_U'
    'Ieca~O_faLQjVfxxb*dU)zbILW^_#<{3<B4=b<RFG_}%WuLR%4YCHC>ZWM1?pF~K}kmK(u_}()Ff}TYTK{aP)p)kn~Fg9#Z%H`epZW#*hb;Cb4W}T9P'
    'XV4mvEHgDW`+{TaOrOxC_d_+d!fw{<y^m}>v!lq6D-|<Zno%T}4L>;5Yx)yZzWX5wB|EnqERUS5N3h*)ay{QoLn)fUtvtk^N0_lL_T!63`(;EDxopS!'
    'wOp^jJyYzr2hXuF{xM>TQh*!lU&n_$fEBLrn^GlMW0Mvu+QEgXXY@(9!qf2!n6+)_&=9)3fEkz7z;`gE=Ad~ywh%&o%DMBw?Z6z_Zsb^pMxw?CzAC<E'
    'k|AKnP&wB7^1Rax&pR>pUElkv(fzmex5Ca0clsj6_the3MID&w`mkwleE{xcwwqEljN&0&hVG(@6_C1GR;^KO6Gk=F7<Ba9Lrt(22xLZdwpuIAXl4xU'
    'cJ3ABS&qXmEV*i5xFs9Jac06ieK~gWMbCl8um>qZus{OS@f$duSSe>6f7fO9MuFwghXB|4LWfbA1_MWy-!k>pU;rejuh}krvt7TRtwz-O3aivXSZ=m{'
    '@7nDaVyyPCt)$|{HZ_!ruid^*W#UzS*36e1@V7#w1kM$hA%W1?nXR?CQZmDn7-$Q4Vhpbn8{9d`Ld;I#FX2NIveyJw-+?CNi<$tU3ycTarPz3@=2o)+'
    'R<n}3A1>Gju{qND0p~goi8i@q!g26{aV^D6>!QiwPb`9OGL}Zm_L*1~ZAeZk@<fT?(n{ng_HoQgtPIb`Y)KIf&7h!BENF8<UlVNDo-`$nyvWVT_0H3Q'
    '3@n?<ayRq+C@BRt(p3x=hOTCFKZ$D{mX*Bha@8@0hstuGf*AURee!v><zdP>>`)=)+~%z+$*tznl%X`(VI$Z+&NsKgSko?w-y)&NA#aiS`tZ3-50s3^'
    '8!zmY67XX5=yH8MTd?0Ye*cD+o>6tSbL`iyZ~L{|%&HN8g1|^9NVv{MrjusfJBlJr1P1#9O+;rdfV$ITi}1Cazdk<y{^oVRNL|J3qqWo~2fK6s2RtGx'
    'cR;>tm%-sRf)|q{k_x+Znd2p8`p%(fDz9mxtG0u}Mfc8v2tmG5qihZmYPa2aPTJ#abvxUhdAaC*;~e5DFgxsK`{IX1yYbv}?F6{aR<LjX;BaK|<>TW&'
    'PG4Uay8N;FQL66}JaV5t7f(-*zn|P+pWGkOy9{+IY=c$z;J(D@Dz*Tj*uLIeBqC!9(xh<I^lrOe`iQfe*?vDSAd^fq3NQHGaRmLYWBjP}DU>>I01Wwi'
    'gaT&j!{XMTLy3m&-ps8n^E#<-6>lN<X9~KDs8NZAgNQmA$vEILwNXZQAQeNZKLHt%f+C_w5e*gL3O`?p$kD_IR5MWhXMm#i3Y{~x#)t{{7E&loqyk(k'
    '{=_TqlI+4)9=Y>z1lN!cKL=pnQXfN}6;S^!XN$M11LP8788TmitZ}OL>%8au{cYKABAumi20!o|rlzTwa2#e$!x4Utf$kF2Fi3tpe7Q6Xi$X%C#!_)^'
    'YnpiId)SR<_Jcx-3n7q?(j6Sa>2XdOs&w46Wr=6f9zqLt?$|yoLUu?JweTI~rIlKiqE@JFmrWISw0&V>t3F2Y;AZZRJ|xEd14w(E0rfx)K+hh2mgN<c'
    '1r5+nCy?5`k40pNbIFiB9LyC!XQDJesC7x>8+Yn`QRDEkGjp15IMciN&Knng4%=8P6N1Db2g$}@%m7+qXDotF+P&*g(g&qzDTSY%$MLgX#m{=hbH$Gs'
    '*lg$|^>Y)Pv#&ew4Z}Bo0zQIrD<Sq48>zdENNz-qH1alEX0ZVb%*Nt=Rn~pO$mWH}$zBjK1ua(~m4scODJdAP!6g7UF~Cl-^_N{&PV<4h;VYovo&mwV'
    '`trdCQC?3)3`GM#ADS&KR_u0XN+}aSm?;v-$ln9yd+R39{;)s1E^anX<|^hTfZM|zRKl1BFISzpO|yqCdX7M3<<<tojY4v)@1XC}Ey#x0PN!ojT><5m'
    'z<^#`hnzSMwq4?7UbKc@?dR`iTYz1(NBAOs;do~HhubSrra1zE&yIfrBHV<55;2?JJ-m;aA`NxL^;6x}Q)eO_^1KI6vcC>P442X!)kBvW{Vi4HZ(eqf'
    'zZKKd)9Lre=Tq@NiplfoAD_R%)&0{6LO(W8AolSW_W|QywE6I>d&%cN{`aTH&;Pv;`nxdSh5hjBn_`1|Rk)6}F}N$dCVqx#yMa-^^)4MZe3R|rU7Tn?'
    'Llxc}BNz_>L!HIg%^_3-EMN+;Qf^{4LTP~kHE!3H+>>qa0RefUzYS(eE;o9M$(574O4xNJ*&_JU!>pYs%=9p?;LFEYN1qG?q52#Wn&>~vWofTUoqlMy'
    '6{c+Xq#91wKTn)Fpm7CUPu+7@MEvVVC;0XGC&>Azo9$}u10^9eX1(pqFPx5sr=v+kA^0PolFsbVwr<?<9HQHORIN}+i$MeQJS`SG3_;T^h;aObt{Z{A'
    'YiX3ESq)i4^c~u3h`xF2)|y+n!H+nDxHqT&KvLrpqkr9u@Yl#t!WqJl7v0<|Z8XJ3=!&**XbR%Ici}!gu(ua=`CZ!YqKxc+nAgK=Rx6f|R6(!e7$s?Z'
    'VA8n-N71Y3whm%lhEpKA2PY@|^MSeV?<G94fk2|o`)Ll1HKhT%9))(w_~HKG%t)tFY&`+%e>&47MZR0Cb7*9&G{&T#wg(NljPrdLZ8J|}Z!7Tg;jOu-'
    'NI`ID2k-;M*iKdW2dSEahHy?naTrZyX&(;+_xrhP-oMfA?d@U??!MhyH$|dxNvq07+UQto83L|V@<Q&+gvd790oLgbAl>QW1dC~`0pRu|XXZ9eE7r6P'
    'xL?$*`R>LYq#3Q>1PIe!g^lVV>|5bT_*v*HXPSkNy4yIFt0Ze&lI_D>c1dR;1k7i+Y})y*4OZ~bs^&zr=#||2o4D0}J@j~R?(k3lI4!jS0mZupDHVRz'
    'Z?^8B1wN_4W@>H*C+*F4c76SjpM1e45Fty&D=)P!Ny58_XAsqMRR~o}cf!iIH&TM-0;+Lbi0%jn{;I~}&AY52S{Gpsv+a9NaUV*sxMRG=Eel4q0-YWE'
    ';|^X6(cv(C44-FRi`v)ld1cXL@5ma;u6I*Ol^yv$a~55*zmf#!9SabVyz%5Di~0KXbbbBCgNOGmD>r*W2^c;%X`n+}Kdsv&WjJJmmKA04$)?*i%rx!1'
    'Vvxu%DCS?vx*5bE{U1XTtkdTBWr_eo698VUQT}GK3VbS{t=v9+{_E%Bc=~XHujNOE&giPFkjfi>`d%E`o$I<Ezb}s;{_b?0A3vO)AAdW&&L$}odpGEN'
    '!k+-a>lzu87f@X1dYvynAl@*+*re#dhD}7qK{Fmy*{I9Fg~O!n9<-uj0$5ey8{N&_Je@hqF-}(;;c(JA8BTw4rl*4;yRmBM{Cnp^ejj}L<qq=?@`GL|'
    'q8^rlwDM)^#^7!Dx<z)6d%;W$zJLW-uNAQbTSZ7<T26h*iC}D!qaZD10Tq8lStc=_jk6mYNjwuxV+ipt2hiu<!3EfZd$;epcy;>q^!)MZ>GRcJpH81X'
    'e7^eh__ycJSI?jS{d1NsWJ-=RGm3i|?mxWs@gvy~$8e$WnzN|Hw>LgfCk7KH!>DT>MxDdPOSV@LkhDoClGVO0Xh=1xebH)CFh?z^j7sLfKF}nXD8VWG'
    'Dj;743kra5HVy;BNwLGiG!3UsmAEGaYw001n%mepAg(RZUjZ6CFgw*HUK&_cA&BeEyI3>F>)&9Ay$#5l+2R24ojw_-rnDmOc68$=@KqaS@svW5Ti}p*'
    '70{CL3%IXyByzns%9;#mF<U!*=K2zvo7T=otSUAUwQfkuL$D+0Ort>Xp37zBj97tc23(nG*a;PTL&=A`nY$(7xg?HGBk<|B;98sAuQ@TsG{%`vNaQS&'
    '0}r*;FQ;d>30|Gt3ixvR<G;Q=KE+i~+ZyaPC3HoJ$c?N;S6-)wWINL_NqckG;urF??-lNyL2vYR!2G?7GoDwW01T*EQFU~G+9f<PK@+7EMt6ow(S@V>'
    'ApVXrXrwz#Lv-l23qIZ4-FDH1Dk&MFwsjwU?eA@%-iJe?yov;Jxyk!0C=x#q1*M3amYB#MxaLZPb1eX1j)%NvKqg~4ctHbq3hxF|>JS)4RVOd*6!IKX'
    'Xv#oq%I=u@ifB|(sxy4cy9`EANhK{+ckm()!>zZ);F8Vfmbm@g?Ma>Yo#aN1qEZlA>!uI7UvD;W!vcJ&Z~|}I>+P%qxz9g|wvz3n?k9=V|Ec3Q7Tig^'
    '!S0Kv_RW5Znx&gPHPqLrYt}6Blm`Mor%@JFrcO4xqoPYz(Pbu?suUD|qZ^L;ENN9brD_ht0-nc%4-fwr^g-;%eGJ?Y;Gi|^#K+EtE3DhYe7kBF*DV?G'
    'J2%{J6LhbAa7!Q*a3Y5z>8A?vt;wmc%GX(wrcRn{{#29R2l+sb%5a@zRob46KGeE_^TsV{WeoZytR{DcGc{N9UGTxY!oz!c_~U=hoQ1Ou1tf3YDYQUs'
    'W&JTx$Fg~FwrWzNkDy=p%j*@mHP;a7M2Z+5ui+rSa;JB20?rpxfV`45wxOMMQ)y=vP=>XU26yUsPJ~LG%Q^5<_;mzj2KvTALi^IbO$(&*sb_$MzCzW;'
    'oex#hZ3@wsZ~-x$Pv*rqtowdffOl2Nt_#;0#FBl&Ro()6#@vSl!!^8Jsy2ej9fM!ou`D42yEGxD!=HoXX1;ezYYO|multc?Ld0T-6yySz(;^c1M~wO+'
    'Mp*f2KV)YjmrZjKBl4c7AaEO7^L?Z(FM*)F6sI_3I~up637gy;Rs>?=b<d270X`t{*Dl;->Uilxo{{Jd@tzTBaSsZaEqG(&cFJAU6XAKcsb@R`x^75O'
    '48eEv#UwD<2c<Zq)Q1dqWmIxCVRXay)J-B9Q^eg?uw1{Jht=-3k7(DM#)j@Yi-?q@sQgBNNbiQ^lYZsj7eXG~F6p}A%ph<T)LNQM|1XIV)1BaT+9x!S'
    'BKF&9(A6<WDyG>7@R+6VH);{@uzXlj2mquiE0|K$sRr|1mBp`5Z;I{1_R9VAHEw}$HT2Z&Buo((LlQl~sLzJ?p5k~L^nk30MUY7^a}{sODn*DY-fHe|'
    'MIL^x%i51rDWrUz<!d7I8>IVt2LhwuP~AG~&WHFW%s~l&Q4CO@2k=*CU9WD#g$s1h6iI#uxn4In%;xB}E@yOh<s+TB_X?}9AEG?$@$k<d_Uk2>=07`}'
    '>mRz8DUBs@qq^Nk@!NKJcn{nd^Ii7pS~vXaoo^VZ*wgD~r&>t78bf}c-Fo9G2=Y94_ol0V{vR77Jz=36;JGdhoiZ{6Ebi9Z2>@VauH<Th1|OIMZZ;sC'
    'WE*0=;ku!E!{H_=0*;m(R;F~Po<({hXL*I#K%hqa$?sa;5fD@{3b#%vVgf`N^M3<GH44F*Dt!zQSy%%`J&NJIU9_NvIwPBkfou<mV{DS(^e&9&L!^B2'
    'Ix?oh7x)Qu-I;$Ae}aZGL9f1m3*+x6h@Nyi+Uo%lNu#)EkY6K%Xi!3^2R4+srUhlLHKS-RPhp$gLbfDTE(!N31CYeG+zmEs1kp9?z~eV09DT%R2g?t='
    '5QYjphpdGmxLDy0@;+y)b%raeftWa_&sE&bW(2p(4Slw^0o8a7tBRpL1Be-8^yS&&3=qNm(PQJ8vq`78LS}WYY`EIfMvuWQ+nrYloe>Ae{7YUO@6y~7'
    'v12bE19b_m3126Nq6;pq^*EFd-M!p1Ce?Cq)z}0|VsmrI?$}i^Fw7Iq*s<lt(Yo1(=udFBTjNz>q+S{V3?aNPm@N%v=UW@nwuaCG`|D*i6KRk`3~V)m'
    'mIlMV>c6H>N7tgZ7>}YO)Zz93u}hv-2U6iN6A&+aoJLre%*o=OH)Z7Kfm@WYbLJqP)G|wk2d{Cgl2ZA|o?*eXF{Tui6%|p422O}6^%P|94$DbL8GzjW'
    '(3CQ;b;0mCGs-EG5NZ6TaKC*z&+VbzUPA@HO<Rb!l3nKkBxc(N$`pU#Tn14LqJhS%g2b%2CS_ol*|$p{Ot*~}&C|D`I3bZ6fghygXgUE&DZtr+YukoG'
    '<&iATYd1(22RVsraH%RfKlgV{gX9-LdE9tW?+#ppYoXW(UAzTikK@`0>^ig?po<`B^jAs<I_pWx(ml^y8PccU)D%q8G&YLNf+J~mJzG8i{cy;-Q1qT|'
    't^*e}+4`lM|0KIj>PaF~O{B6^5YVuV>@%?<WeD})pZj*dU;6+LuteGgU!Xipucji^`;DMrz3|h?9^j&6KZ~*x(0edmoec>NToX?HYtYH#lmlH!;r_2;'
    'TyYr}OsA<Q0wzrRy3<d|c=82!pjXYWxrb;dx(TL$I<$;cheMcRJ%q?Nb5JM}k!dV_6r7ZJB~ckjeRBJbm5ghKaC=Aq@6B<n1HPO5^Z!C=HPY~P_%!c<'
    'MHNE3z%?6DLPCKT`|a`R<ZrW1*n^OR>%)z%NvDWcOO(q)vY$@Rr%(UR!v5&bgd@@Oko4o<o|A^l0a>3p)s<vdrXbr7-WjDm(}oXa>pq^8@e+YT4IU5e'
    '+|F-(wsL4s7R!VR#%Zk9#}+fMz5ht=<M^uf{o15E=2YEZ>r#?KT9`1nE6->YSH^aBv);as>F{iGxvy(W&x=n%@_uV_OLV=Q?ICP!eme*6SBPGM<6*v<'
    'E)EY6Qsl<r70}<SxF#mclHwOc6D-FOC}(ca{n=uo@CHrU+SwQqn3*s$%ZB=+`+2*?)G^Q-3*wIMrmlobJ}4vSM`nXnIwBc!O;EJcBjY`=)u=Sm%~<QX'
    '5Cj>dbs~;Z@ow=R;uve4Mf70lEbv>!?=ts#G%sS%K(5zs!T$t7p2xY5kX)l9OA3nY@zHml+)!vDW|S&H1tahb-t`njuS&_mLUkPQvcl*ma6B4@+#F_G'
    '&(jfV7ouZ5v;q<6{xTWvRUHKS-1dv>3GMy<Iu!zKy?YIwk1Jg)_q&W5sNytRDNyd(d&a^<-Vw!?1nzc)#m?Nkul%5ZLvn~-YcFL>G{Ih_0LD6aBlru='
    'hWM;Wp4FXaP5;>mKP$V>PV#4)1h?uu+w^do1lM$*&2d%wZe{wc>b;wi39jfqo8cymXOuRsLM7%sjPH~#a~r2=xgASL(riOS%6>arfnUi@qAj%XlMlti'
    'h?4EWDIrAtUPX4ky_=>bEn2sU`mJs6j!o}$yq>#VXpQ%$#ogWyI}>wxU#2aYn9Glyr>z1Sow=LSukc^ruc_*^6Rv}w5vv0LDO31Jhl#&u45aaJGHzx-'
    'xEl$UGp;qu^?^c}8mB#MJufzrts%#htWZ1jPw0Cr<?a+Vok%X*)!Zj2cqZb%J!9Es4RxEL?7boaInePn+?+<$&zbf<iUyJ9yPqDy$p>g249ZS`NJd5*'
    'tDE{l-u?vi;XmoVg(sljS{~l{+E!`Yo#wbZjq;26ZcV=KxN&DUh`1AIF0|mU5xl&2b9pMY5OEvt4Iy{{ie8iR5GGKALVK!82t9o_Dd55MKFHZnw&ak0'
    'f7QT}LDLP6dErC(%d4{LMhH8rBIxJZC*8Q8$J2j&dwl*}>}h%#XM`wWL`{_Ip)|-5juR|h))v|sBTG8na$&`nxZW7bh3g^VSOdf$4-<^C2*UH34@V)Q'
    'vAtWK(*>Nbk@fm(4#He&(i<2IvUGi!9v&`(NKS2s<kb#HAQ0UEmm|>HF%<Ymkl9*7A+#<5nZ)czDC~~?@_`GhIfS6VxfKr)pMBuJ>W-+#urw53uMh#Q'
    'sY6h240)o<qVIw)6^-9bLrHnBDwtG+e5+Zdv6Ryq&iY1Rj!0mmt{^bm(hZ1F4Pz2F@$7!jH2usntZO$Byp>T}wu4F^?PTK(w|53hZ;T#K&mVs~6^n1b'
    'e*E<L>#4Z;_}AxC@yGxA^4rG`r()u@cW2MWSo5f-3E*HsmyKIf$t(M~yH1|M{<ew)c>7?+676C9>5D74GyTH{t^hR&Sy+b+tF3!K_o)02{!FFcAGjhm'
    'KHl)-P51GnXnR(nK=3hQYLQ7%VX3O~Qf#KntXpWJ6uUQpi#pOMVGj-W&}5QWpsUkT#Vu9c=b`>gHgri_AYfS5<g&0&o)vCbI2+2aRSv!(JYBP3JtXL6'
    '?t8dzA8h@2)CSc<w<hcSGon&K6Xn=~tB#r!^f}t)46J$xP==YptvQ|D0S$D#<c^np6n`k4L7}^-N?`w`R{YVJo8E~QAS;rIK~8V{>1DXN@b=mApeTL('
    'nF%=+!?zPhYGR4pSpijhUqC4kNF??ei~e|y-l`gt0c&k>0CH__uy!Ur=_&X#LzuQVGYF^cG~JeBa!r{U$&f(gBGCA(lN&X6-g_x`maI3D9EskU?xoxx'
    'HsHND&YYh5wz!!Ai#dD<r=YnbC%gTUWJ{{b(j104&!teq<=0@%HT0lDm|!^~?imq>>t||W{%EX{!{Xh0n!JQ(WGzt?IV_bO#)<4DgO%er*fIDb2gK95'
    'oI~$YD!e>!E3-5S2xAOVUIwr?`M{F9pbC+^eKdcI*sBZ|Slt@tVY9=1$vEA{YUNfn72j66^NA9umzD;12fkw3Iy-{K#BOGQ4J*{jW35rCW*8<C8D*mz'
    'sa%k9-kMNoxG&#f%Nd$AHZZrjJ-i$Q$nMz&C!T36my8<z%E)6g^qD1-f>l-wi=Qv?jr}OObd?<LN6K(B(3hqp+Orb%rXIw(%IsMBk;2xi<Y>xZ&X$2x'
    '#Nl2A(w#oSP))_SA68UFNLgzKr7u)V<XOfAi7KTAp?;q4?o!!(C^%=y#n$lZz1g;#O;CvUkj@s)9*m8iI-NgfaxwH#t)W*rSM63dXv(>zZfJQX8yKh;'
    '7@ZdAuACSgyv-pGT@WRc?lyg>vr!8U)T<3iA1iGvaJ@IE!U{K22XH%FTp_7`j(GP)Q~0~|w2%P?ZWum7DmjFULCFK1qv_L7OmZxzl%aT|JCdg0$PK_)'
    'I#%7-i4mB1VbwzkmGyLEF^N;eyVV?ox)EO|8ye}(5r}pu$v{?t88O58Y1`iHiOjGw{-fB;rBrL?QW^mF{*~4<A`g#yfRUsCG-l{sJ<M{w^H)1(fTFJ?'
    'mT|56KuI0fGcZoC7tAoG`Y@4Y?95cEsBDv)_onVc(C074zCXr^!S^4SL}*`}`$%D1sD$Av@q`4?YuAZt8fKzfV>t~O8fQR5Mq_~lWNs6)?$n{u@NOZx'
    'F{~RZ8|GF(MXwK+FQF#ApY^hezVJmU75U9WsBd_F3ub$$(E(Kp=5XIrY!d66jbk>c;llAOf&)OC{M(xhri#4UQ1w`dOEAk8Or>@Z$zakIROK+2ki5}0'
    '_+FTp3GY>1f=R3!?$hz-yZJqrmths#{daJ{lJB6X!Z}b`H+^g1+p-~e$fsKPpWwT4D|)?b_Yknqa=JEoSLiG(A<~5^QHNCk;*;6x7RX3qPWp8k1m|90'
    ';gsyHf5inP@rA3r_CD-1;D`#kR*DopXQA;~LC@NeXM;&QnFLq}>yTcn>^vGTh6s3)0-Ld?5qiyC4jK8~r`H|iHA4%3QQ3Vsx}Vkw{xlsCnM*M3uExic'
    'zSpCVM{xI_U1O6iu7%R{UKfsv+c_9_`)$-wCq9~X9ZmA>{$7xL)4V_)>Vt(?8k$40w2G21!p{y0O`%|&`Fqekb|4ud;p-}NM1vj~hjt3F`zxQ78Ar^^'
    '^GB)VOXcGbZj;8#k;dukY?GGCk(LN}oqqYU_b_<d7&x6+VgK$wEa-*rXf^n#hV0p0DBil5(en#_z8QYr?NKx#ZrAj<C<h;hM48y}+hqLuOEXEUp;x_4'
    'UbR0{digz*EiV5cAqOQGK3)ECNAEu!q1(>l?|uD!xVnp*(b}^#Ym!WM0{H?-Y%-Z};llBZV0V(Vo$04vB8r>PB%O<PR1)_W3<#%PDIL?Yda-<L{yJ^m'
    'jt_H;63AP>?zC=RDjpkO9MgBNN2M1Kn8NO0P0~r*FO^15te1)m+w^l70R8+y<HXLd(&nYQEZyzFWYrn@@o5_F3i5$pbq2n9p)@wVO4gRt7liQzxb0Q?'
    '=G9BZW8>>p`rhM>!^({JB-rpeZMd{A5si&*(l=J)OFNKnX3cKWW|erEFgBbXHCA1a?d)TlbRQGtOGIf~PWR67aSHbw-I;aVrZcN5UnYzV-=qt=IV0<K'
    '`e=Xi@o<yAuXvd-9lYsanex2wnh(_Ul&S1RG4Ad0CY@2^k}=pId)H(7uDX1obR6Ylsu;=fQKqPz&QWztk1A<jC>@*Lr!U>#F%R3n9N^l88r!EUrx|}9'
    '-)x!otv3a_mx|vDmP*E>yd;I=kj^$WEW}MDEeJXGFsi+7=txZ)=ua5f2t4o1`g-JHEE*$ND(R59Acb!{Pn)X7gpyX$M)dgex*K}fG?cdGj6C{^ksl8-'
    '^n?N?X$Z0hA`;mZm+^K3oJWZomWWf%`*#I#is1U-2<}itCixkpV_Mx!JDvHu0-}V#nv{IiR#v?309gjI7l6Vhj5rd*Gb*hrXS@OT1!6U9Lm-wn;;@jy'
    'rd<F9pYTyb_k}+Aedld#!Iv7ZWrt=}GoQ0&AHe`qF;NavkZZ5O@CHg53S)<=yg)XglSZnSH4?f!+J=7J<*lmAl6b2utFqvuA>RN6`O2E75W|!KMiBi~'
    'ZMraiHyZ?jtL3{Eq>)T027}oe$uS`YykQ-Eby%kLc_y)5^y;7gugv8e3OR3hMvu$)+Na8JPF`G=mzlgMx&N$cj?1fDUL^fjm*sUXFO&Z3%kot&&oTH7'
    '`=iF?O)gKxjYPw4r*Zi@m#+lN*O%p+T%JnDiNMozLEdJCqeR91Z}Rr9*WkMBukETVIyHL{tLDBm<`NbK<?BH-RZ%v%up(a>#H9dn;;W)8K79UE=;HUs'
    'r;op#zUD|(MM$zdk0j0`8OccEJd!+*MAhLOl01*3&LdGof<sd0k@R^a3P=!E>hnm|StL^?Z&#g1GUt&*@^<DtQhgptCT~}tM{3R^QC&`9Wpf_Mjv)Ef'
    'TiH6<z4AHkErao`zp@Q&UA8PQ2#Z{pibFy)^abI3(~r!HL0Dc8&UgH%umoXsK{(&;ODR}bUl7js`cm<ZQdSp)^NpTa#+|UaAe`^=WySkOSzi#&xA@ZV'
    'Vk(;p!Zrz?hFnzJAQ3=%%M(gvfQIR`lwf1Dw+DY$R<zsw8wq7_J3tUcjsdQk?Gyrn<j!caGm4TjM`3bH&3b<aC5f>uMqnx?D1!mYVI;g&>QY7^X~NhZ'
    'H+cJ3SG`VTf!A2H`M6nFG`$kjm=r$m3yjrSSO5G^8-+F{3YF|=;e0YY1LM~$QA)6sxF{u1O0tx^C?!!!v6Q+frBF(<l)fmXQL198>Y`MIQii3>MJa<)'
    'HA~eOrD~LFSgN@wh0|$cS;}6N!nv_6$)jcZb0b_V!M;zvC?|0*SV&(G!r8Rt+eIVBMMTn$=(JOv7ov6~w<E_zWYUg|?SxC-3qsV6;&#-yh)UW~>H8>q'
    'UWnRZkE7A!B06bDciJ)Mg{Ym1+c`7TRnkt?X{R|aMC}-E=S<g3(vIo0Q=b>2c4}_tj7XidQ+MW@IWI)*G~CXaA!w3znoc|Qc_C`Yayw_Hj!oLJopvq='
    'Q9DyUEcC#{<~!c}yx30pz|i_7lGM{F9~wjJye}z9k`IodFV>e-B*}-z&^PEyYLeswL^m&ytVog%k)iL=moy~F2T4`FM6xDHK1?d{63K=n`9P`UOC;l&'
    'vSockn30ddN=k`BG<yOK)mb5dLXw5VB_WAIin~`|5>hCnS;$@z!n-C`HH2v{2~{XWlSJ5yLI#Cu7P>g(H3~H>ba5aW6k-juR+ofO%Q*6k7&!QX^52^h'
    'n<Gz&0W~9+!rkP^b3%)YQn-^Gc~WS3Q3`jFBhQL~%`%t59puQ<Vql}prEvE+^1K+>By%a;IgUIr1~$lC3U`eo&y0aBE|<a`<H%D(U;fHrZkWm)Pg;bK'
    'SPmtm57&$EKobc9Sv@%`2FZs`?D52tY~wOUw5Paa;_|SZC)I_5Ujj4#z}Tg2Wk(jSN;ATNX5{Up-MN3eDke`Kzy9|442f8}_?Dr@LVy8P^-$yc?OqJm'
    'umgz9Plnt2qSBYJLKcv-WCazqC+%-n+{{-#6k61UGiF!7hdx=%J!_}~>5O|ac#2&sw5z084k}4ox$u^?VULGZ!)W+<IbHe^XvIDl^4E`o(ucW(s@W}M'
    '%_Y<r8GI4dGM$g5rRY%C+oc!wtlzVH<QMRq$Km=A$qk5CIr+&Tk(fEmXNpy1o@NAKU4U+g&H}9Atvd*;!8G+n1iWwSyoF^2(~64_cfi>Wf9UjrHL%89'
    'f^e&--FL@9veqc|xI9s%`rWS60ahNHY8-&xe*nU=woJEO0}yei9r_lQb$q(78%Mx_OhU`BoF4t`A>?ec9bvXL>uOv$1Wd-`?`Aq>p~1J_0?k|2pcC7&'
    'm>m`CHZ_~D%)s!_U2q7uCaI)@%^9g)7iZV(eNblvu__G7C9fvV$i?X3?k8aLxUSF1b_V7yX>p*Zexmebl3~#!GM=ZRJ1o2dDt3A|h7sn=4Y1hqrayg~'
    'WN|_G7SLu4ZnI5qmsRZEZj3qG2CqzrcAfdkphG7_(`$qM?@e;Z_S%SL8Nvb1Z;+ws<jy;pq=be9Si8tPo}6P3*mI{HJiQs}s!lG%lc`1EF*YVS#U4oE'
    'z$N7eqXUl0MpcutQSsdK0Svns@jr>$L74L0ffZxP%+4NBjBL*%qd)Ye*pVwGs7Nx$<}2kL^%+&Arl%<(q>>D&&Ovf#FN0QwLG?LMzML2b4fDZj$(QQM'
    'cJKN9#xx{R=Bvy`YPtYZ_h5!+y49N3O|r3)6zP%|JD?lQ^IGzVC%RtLmvu3O*wdViKt35&no97L+yCrFG*xqxZb>)AJj(B19-lw_{o}uV`+zjr4SES8'
    'B@LKO;9}v6A*6-MZrB>)PE{H&Lq{=o?P9rJ#S60jEt{}a&0zKdSQ6w&H(Wi$2FYRoQC~oeo6FEOVC32gRNaAV3r%64dc9g7=j~8Q1Sg{`%a}ywI6q)8'
    'cVJW!Koq5c?KEsBp+<6pBC!picT0-EA#bbov{#KrUN6}8qvR^NmOs2l&Y-q<rN0+&=Xm<^{Q3KvtQJJYv01W-L$^wD4cxcfw<}*R=-%F|DwF2O94BpA'
    '+8g5U+M=C8((E;4T;2=sFHYj_1c$_h%d1ds1~@@iG45@)SPvN0J^)VY7yx%{VTk-l*2p7>P#dVz5g0R)aXe0-Skv8%Wi-op+LZW!ua|UK`i1MpGzym$'
    'vYtS39+8M1Zu|#3tTWD3ThkU3PU$Exg(Gw|4~6uK<Kg<yEp(wPviWo<r-Oob?H2Y#Uq)_mxSw^ui`&xBc$YS;`JUVYs43nZ{B{Ck{ciox=B<+*tK*r}'
    'L)k#l^zEz-2E=h%#at3UIJ!!kc29Q-2*Oqqq;4TwS+bNJk^(lg=z^(Y{B3wVTr1{}@8`GM8CDaL71+5(P({pIg{Taa3c5Yjf-IIxqp@@bmBbPuR5-FC'
    'mOXi*K=q8cL`deh)K{ULOw(k(AcMq?3NqM|X*fMV0lQtGcT5P*e9TR^o=II&#YbC7*~hrQhlEqN6REmWrMIj?Ns$<&n@Dzwc~h0;gyL}Ov_B60HFbP1'
    '>DEI7p&%uZ;)ptix<_PupSQ?d`P~#1R;4{&S5(qgN4m*WaNXf;tA)Y*UU(OgKiv7Ee-YIcCpQP$5vO;MmK&<hqP!2}+qKAN><Mh$O!K*<7#*vFHTBUm'
    '<=;F$ef&o;{r2mZr^nMD|MmH6v3&gP<o?^~ms9cO>Ga#<>6gzbQ#Nrl$WkS9>Bec1ZL!{jcnh3EzC7u2DWt*Z+T#uE(7b~sAI}r&#DL#q<8t%Zy@Rx+'
    'j=?sbA*89hf{}v(My(qp5|`@i*m9#yPxiKO4}=5PDU%x%x+57;^<~z}){8v^g2`?ocMU8S3aa{+ND0z(_OysX)C$F<odEIqs!e?x$D8GXJ6`t1S(kF}'
    'o$mKO1B?5wdZIB!D>Csa;Ek!9IJ10-BiSG0-RwFPq61Q#`E*W6K+fz|GAV4*bjItjM*ft<d-eOjDT%Q{Zlf}NiRdi3Ln7PbO;t8n{KD<hb(x6-P9|6c'
    '@%f#V$od9NG}-gJ`O+Pk>&Q&CX=%o}7l0a6Ry9=IPb9)_Q%4TTgc>8-W-ZhU!6QP~!X`(A>0!`403X1xPgtcr&%MCX?=cZO$abr~dqQW=kOR{YGtr9c'
    'TL4Ii%`4#cRa!FxyBdOY-x;Wcoi2BH2m2j1(cOb$a$3+FQ$5H5t&8@sovkA~`|8%y7kPao8uK|+sTj(<5X#f9qk>uclB&+g<C^E%CKKuon4~>}c2pbu'
    'nJ~*p-|Lc<VU=vF%h~l|*CrLNyn~QR4+2Sg`Y>=pB#ugqF!HRkAV<u3ra@JOej<o3*NZ#at6Z0ig_Pe(2n&u=R)SKkY$P)Z%0jZcu0!$Kd}{`Mp-T4m'
    '=<(s7C=S|*S4h3@7YpBD)!|Scv&1?G5sPaebo9+WmM8nlv0^fcg%@+mf`|mhp1@J3TBR3+3F0P%Cq+gcMmI*&jT6gHPmb>n#rIO2JE%vwTGKajdw5M`'
    'fg(L@y<xT^9)H#(kxk0z^lYbvKA@LmN{GUZ)FLrr1~I#L3Y$OUmcE3!*F~`Uhpbdt=~_~uZgS0JDE)oeO)HIrEC>_5M-Y)|(mk*X<9A*@%LZ^7%+D!J'
    'DDggw1J&-)PdU^ZUGhumF%&9C&Io%?Nv3Nm!5n@LipNh%#F2sJ(v(D*9De4qu@YVPV^+9N%G#}QJuc=yB5>z)M-3T%p%CBDsQKcM{`g5$f$nR&;kJ0r'
    '#-BM#dZ81ZGAv-m1lok;SVEl)ECVOBN);xzPvgEzhKhSUDe@Qj=|e%<K)a}uetL}-Dm1y3vz)DFxzb{&M3hLUsMl6D+u3q{SS7>bh5uzLj`LDYZl`w#'
    '>~zdOj>4fnMX*zg$$Y!NyKWzf+xd3Edpc@A4(VpfWTy2POuv?H0#(Qw(Zq7gXle|NYv<w}h{GrmL#7z!GDfH3u0#UsssXBVV+|Q=PLb#j6T`gv{q+3N'
    '*%2Q;KYf0F{Q8<J16U04z7e`6=I}b*Q@sf~>+T2}U;*knCTdF*kTKpnii7oxB__LG0jZIw+S*Vx?YLe{TR)|C8@PILRznd=pXzU)ektZFccd1}|Msu{'
    '`u_pm+hch'
)

def verified_seed_records():
    return json.loads(zlib.decompress(base64.b85decode(_BARCODE_DATA_B85)).decode('utf-8'))



def money(value):
    raw = str(value or '0').strip().replace(' ', '').replace('$', '')
    if ',' in raw:
        raw = raw.replace('.', '').replace(',', '.')
    try:
        return Decimal(raw)
    except InvalidOperation:
        return Decimal('0')


def display_money(value):
    return '$ ' + (f'{value:,.0f}' if value == value.to_integral_value() else f'{value:,.2f}').replace(',', 'X').replace('.', ',').replace('X', '.')


def load_csv(path):
    result = {}
    with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        reader = csv.reader(f)
        for line, row in enumerate(reader, 1):
            # Exportación Cotyland: columnas 0 a 8 son cabeceras REPETIDAS.
            # Los valores reales son columnas 9 (IdArticulo), 10 (Descripcion), 14 (Precio).
            if len(row) >= 15 and row[4].strip().lower() == 'idarticulo' and row[5].strip().lower() == 'descripcion':
                code, name, price = row[9].strip(), row[10].strip(), money(row[14])
            elif line == 1:
                raise ValueError('El archivo no tiene el formato de exportación Cotyland esperado.')
            else:
                continue
            if not code:
                continue
            if code in result:
                raise ValueError(f'Código duplicado en {os.path.basename(path)}: {code} (línea {line})')
            result[code] = dict(code=code, name=name, price=price)
    if not result:
        raise ValueError('No se pudieron leer productos del archivo.')
    return result


# Configuración local: no se publica en GitHub ni se sincroniza la base SQLite.
DRIVE_DEFAULT = Path(r'G:\Mi unidad\Comparador NUEVO de precios\Listas de Precios')
DRIVE_SETTINGS = Path(os.environ.get('APPDATA', str(Path.home()))) / 'Cotyland_Comparador' / 'listas_drive.json'


def drive_folder_saved():
    try:
        value = json.loads(DRIVE_SETTINGS.read_text(encoding='utf-8')).get('carpeta', '')
        return Path(value) if value else DRIVE_DEFAULT
    except (OSError, ValueError, TypeError, AttributeError):
        return DRIVE_DEFAULT


def save_drive_folder(folder):
    DRIVE_SETTINGS.parent.mkdir(parents=True, exist_ok=True)
    DRIVE_SETTINGS.write_text(json.dumps({'carpeta':str(folder)}, ensure_ascii=False, indent=2), encoding='utf-8')


def find_drive_pair(folder):
    # Windows es case-insensitive; permitir también Linux o entornos de prueba.
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError('La carpeta de Drive no está accesible: '+str(folder))
    entries = {p.name.casefold(): p for p in folder.iterdir() if p.is_file()}
    old = entries.get('viejo.csv')
    new = entries.get('nuevo.csv')
    if not old or not new:
        raise ValueError('En la carpeta deben estar los dos archivos: viejo.csv y nuevo.csv.')
    for p in (old, new):
        if p.stat().st_size < 100:
            raise ValueError(f'El archivo {p.name} parece vacío o incompleto.')
    if old.stat().st_mtime > new.stat().st_mtime + 2:
        raise ValueError('La fecha de viejo.csv es posterior a la de nuevo.csv. Revisá que sean del mismo cambio de precios.')
    return old, new


class CatalogDB:
    # Persistent database: outside program directory, never replaced by updater.
    def __init__(self):
        self.base = Path(os.environ.get('APPDATA', str(Path.home()))) / 'Cotyland_Comparador'
        self.base.mkdir(parents=True, exist_ok=True)
        self.path = self.base / 'catalogo_tienda.sqlite3'
        self.legacy = self.base / 'productos_en_tienda.json'
        self.conn = sqlite3.connect(str(self.path), timeout=10)
        self.conn.execute('PRAGMA journal_mode=WAL')
        self.conn.execute("CREATE TABLE IF NOT EXISTS catalogo (codigo TEXT PRIMARY KEY, descripcion TEXT NOT NULL DEFAULT '', precio TEXT, en_tienda INTEGER NOT NULL DEFAULT 0, visto_en TEXT, actualizado_en TEXT)")
        self.conn.execute('CREATE TABLE IF NOT EXISTS metadata (clave TEXT PRIMARY KEY, valor TEXT)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS decisiones_manuales (codigo TEXT PRIMARY KEY, en_tienda INTEGER NOT NULL)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS precios_manuales (codigo TEXT PRIMARY KEY, precio TEXT NOT NULL)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS codigos_alternativos (alternativo TEXT PRIMARY KEY, codigo TEXT NOT NULL)')
        self.conn.execute('CREATE TABLE IF NOT EXISTS decisiones_impresion (lote TEXT NOT NULL, codigo TEXT NOT NULL, imprimir INTEGER NOT NULL, PRIMARY KEY(lote,codigo))')
        self.conn.execute('CREATE TABLE IF NOT EXISTS preferencias_impresion (codigo TEXT NOT NULL, precio TEXT NOT NULL, imprimir INTEGER NOT NULL, PRIMARY KEY(codigo,precio))')
        self.conn.commit()
        # Import legacy selections only once: repeatedly migrating old JSON would recheck unmarked items.
        if not self.conn.execute("SELECT 1 FROM metadata WHERE clave='legacy_migrated'").fetchone():
            self._migrate_legacy()
            with self.conn:
                self.conn.execute("INSERT OR REPLACE INTO metadata(clave,valor) VALUES ('legacy_migrated','1')")

    def _migrate_legacy(self):
        # ADDITIVE migration: never unset store flags.
        if not self.legacy.is_file(): return
        try:
            values = json.loads(self.legacy.read_text(encoding='utf-8-sig'))
            if not isinstance(values, list): return
            with self.conn:
                for code in values:
                    code = str(code).strip()
                    if code:
                        self.conn.execute('INSERT INTO catalogo (codigo,en_tienda) VALUES (?,1) ON CONFLICT(codigo) DO UPDATE SET en_tienda=1', (code,))
        except (OSError, ValueError, sqlite3.Error):
            pass  # Keep any damaged legacy JSON untouched.

    def store_codes(self):
        return {r[0] for r in self.conn.execute('SELECT codigo FROM catalogo WHERE en_tienda=1')}

    def products(self):
        return [dict(code=r[0],name=r[1],price=money(r[2] or '0'))
                for r in self.conn.execute('SELECT codigo,descripcion,precio FROM catalogo WHERE en_tienda=1 ORDER BY descripcion,codigo')]

    def sync_catalog(self, products):
        # Register ALL products, but NEVER auto-mark products as in store.
        stamp = datetime.now().isoformat(timespec='seconds')
        with self.conn:
            self.conn.executemany('INSERT INTO catalogo (codigo,descripcion,precio,visto_en,actualizado_en) VALUES (?,?,?,?,?) ON CONFLICT(codigo) DO UPDATE SET descripcion=CASE WHEN length(excluded.descripcion)>0 THEN excluded.descripcion ELSE catalogo.descripcion END,precio=CASE WHEN EXISTS(SELECT 1 FROM precios_manuales WHERE codigo=catalogo.codigo) THEN catalogo.precio ELSE excluded.precio END,visto_en=excluded.visto_en,actualizado_en=excluded.actualizado_en',
                ((x['code'],x['name'],str(x['price']),stamp,stamp) for x in products.values()))

    def set_store(self, codes, value, manual=True):
        codes = sorted({str(x).strip() for x in codes if str(x).strip()})
        with self.conn:
            if manual:
                self.conn.executemany('INSERT INTO decisiones_manuales (codigo,en_tienda) VALUES (?,?) ON CONFLICT(codigo) DO UPDATE SET en_tienda=excluded.en_tienda',
                    ((code,int(value)) for code in codes))
            else:
                # Do not override manual decisions even if a subsequent import differs.
                ignored = {row[0] for row in self.conn.execute('SELECT codigo FROM decisiones_manuales')}
                codes = [c for c in codes if c not in ignored]
            self.conn.executemany('INSERT INTO catalogo (codigo,en_tienda) VALUES (?,?) ON CONFLICT(codigo) DO UPDATE SET en_tienda=excluded.en_tienda',
                 ((code,int(value)) for code in codes))
        self.write_legacy()

    def print_decisions(self, lot):
        if not lot:
            return {}
        return {code: bool(value) for code, value in
                self.conn.execute('SELECT codigo,imprimir FROM decisiones_impresion WHERE lote=?', (lot,))}

    def set_print_decisions(self, lot, codes, value):
        if not lot:
            return
        codes = {str(code).strip() for code in codes if str(code).strip()}
        with self.conn:
            self.conn.executemany(
                'INSERT INTO decisiones_impresion(lote,codigo,imprimir) VALUES (?,?,?) '
                'ON CONFLICT(lote,codigo) DO UPDATE SET imprimir=excluded.imprimir',
                ((lot, code, int(bool(value))) for code in codes))

    def print_preferences(self, items):
        """Saved print decisions by article + target price, independent of CSV revisions."""
        if not items:
            return {}
        codes = [item['new']['code'] for item in items]
        placeholders = ','.join('?' for _ in codes)
        saved = self.conn.execute(
            f'SELECT codigo,precio,imprimir FROM preferencias_impresion WHERE codigo IN ({placeholders})', codes).fetchall()
        by_key = {(code, price): bool(flag) for code, price, flag in saved}
        return {item['new']['code']: by_key[(item['new']['code'], str(item['new']['price']))]
                for item in items if (item['new']['code'], str(item['new']['price'])) in by_key}

    def set_print_preferences(self, rows):
        """rows: iterable of (article code, updated price, print flag)."""
        with self.conn:
            self.conn.executemany(
                'INSERT INTO preferencias_impresion(codigo,precio,imprimir) VALUES (?,?,?) '
                'ON CONFLICT(codigo,precio) DO UPDATE SET imprimir=excluded.imprimir',
                [(str(code),str(price),int(bool(enabled))) for code,price,enabled in rows])

    def link_barcode(self, barcode, code):
        barcode, code = str(barcode).strip(), str(code).strip()
        if not barcode or not code:
            raise ValueError('Ingresá un código de barras y un código interno.')
        if not self.get_product(code):
            raise ValueError(f'El código interno {code} no existe en el catálogo. Primero cargalo.')
        existing = self.conn.execute('SELECT codigo FROM codigos_alternativos WHERE alternativo=?', (barcode,)).fetchone()
        if existing and existing[0] != code:
            raise ValueError(f'El código de barras {barcode} ya está vinculado al artículo {existing[0]}.')
        with self.conn:
            self.conn.execute('INSERT INTO codigos_alternativos(alternativo,codigo) VALUES (?,?) '
                              'ON CONFLICT(alternativo) DO UPDATE SET codigo=excluded.codigo', (barcode, code))

    def import_barcodes(self, rows):
        """Validate whole batch first. Never replace a mapping for another item."""
        mapping = {}
        for code, barcode in rows:
            code, barcode = str(code).strip(), str(barcode).strip()
            if not barcode or not code:
                continue
            if barcode in mapping and mapping[barcode] != code:
                raise ValueError(f'Código de barras duplicado: {barcode} para dos artículos.')
            mapping[barcode] = code
        available = {r[0] for r in self.conn.execute('SELECT codigo FROM catalogo')}
        current = dict(self.conn.execute('SELECT alternativo,codigo FROM codigos_alternativos'))
        missing = {c for c in mapping.values() if c not in available}
        if missing:
            raise ValueError(f'Hay {len(missing)} códigos internos que no figuran en la base. Ejemplo: {sorted(missing)[0]}')
        conflict = [(b,c,current[b]) for b,c in mapping.items() if b in current and current[b] != c]
        if conflict:
            b,c,old = conflict[0]
            raise ValueError(f'Conflicto: el código de barras {b} ya está asignado a {old}, no a {c}.')
        with self.conn:
            self.conn.executemany('INSERT INTO codigos_alternativos(alternativo,codigo) VALUES (?,?) '
                                  'ON CONFLICT(alternativo) DO UPDATE SET codigo=excluded.codigo', mapping.items())
        return len(mapping)

    def seed_verified_barcodes(self, records):
        """Non-destructive bootstrap: add trusted mappings, never change existing assignments.

        Existing catalog names, prices and in-store flags are preserved. Conflicts skipped,
        not silently overridden. No fabricated barcode or SKU is introduced.
        """
        current = dict(self.conn.execute('SELECT alternativo,codigo FROM codigos_alternativos'))
        known = {r[0] for r in self.conn.execute('SELECT codigo FROM catalogo')}
        staged, add_products, conflicts = [], [], []
        seen = {}
        already = 0
        for barcode, code, name in records:
            barcode, code, name = str(barcode).strip(), str(code).strip(), str(name).strip()
            if not barcode or not code:
                continue
            if barcode in seen and seen[barcode] != code:
                conflicts.append((barcode, code, seen[barcode])); continue
            seen[barcode] = code
            if barcode in current:
                if current[barcode] == code:
                    already += 1
                else:
                    conflicts.append((barcode, code, current[barcode]))
                continue
            if code not in known:
                # Actual IdArticulo/descripcion as transcribed from user's Excel.
                # Price unknown: leave it blank; never fabricate a price.
                add_products.append((code,name))
                known.add(code)
            staged.append((barcode,code))
            current[barcode] = code
        with self.conn:
            self.conn.executemany(
                'INSERT OR IGNORE INTO catalogo(codigo,descripcion,en_tienda) VALUES (?,?,0)', add_products)
            self.conn.executemany(
                'INSERT OR IGNORE INTO codigos_alternativos(alternativo,codigo) VALUES (?,?)', staged)
        return {'added':len(staged),'already':already,'conflicts':conflicts,'products_added':len(add_products)}

    def write_legacy(self):
        # Maintain backward-compatible JSON and prevent lost selections after older reinstall.
        codes = sorted(self.store_codes())
        temp = self.legacy.with_suffix('.json.tmp')
        temp.write_text(json.dumps(codes, ensure_ascii=False), encoding='utf-8')
        os.replace(temp, self.legacy)

    def resolve_verified_barcode(self, barcode):
        """Scanner: ONLY an explicit barcode→IdArticulo association, never a SKU guess."""
        barcode = str(barcode).strip()
        if not barcode:
            return None
        row = self.conn.execute(
            'SELECT c.codigo, c.descripcion, c.precio, c.en_tienda '
            'FROM codigos_alternativos b JOIN catalogo c ON c.codigo=b.codigo '
            'WHERE b.alternativo=?', (barcode,)).fetchone()
        return dict(code=row[0], name=row[1], price=money(row[2] or '0'), store=bool(row[3])) if row else None

    def find_products(self, term, limit=75):
        term = str(term).strip()
        if not term: return []
        # Scanner sends the same data as a keyboard, normally followed by Enter.
        norm = '%' + term.replace('\\','\\\\').replace('%','\\%').replace('_','\\_') + '%'
        return self.conn.execute(
            "SELECT catalogo.codigo,catalogo.descripcion,catalogo.precio,catalogo.en_tienda FROM catalogo "
            "LEFT JOIN codigos_alternativos a ON a.codigo=catalogo.codigo "
            "WHERE catalogo.codigo LIKE ? ESCAPE '\\' OR catalogo.descripcion LIKE ? ESCAPE '\\' "
            "OR a.alternativo LIKE ? ESCAPE '\\' "
            "ORDER BY CASE WHEN catalogo.codigo=? OR a.alternativo=? THEN 0 ELSE 1 END, catalogo.descripcion LIMIT ?",
            (norm,norm,norm,term,term,limit)).fetchall()

    def save_product(self, code, name, price, in_store=True, barcode=''):
        code = str(code).strip()
        if not code: raise ValueError('El código no puede estar vacío.')
        price = money(price)
        if price < 0: raise ValueError('El precio debe ser mayor o igual a cero.')
        if barcode and barcode.strip() != code:
            prior = self.conn.execute('SELECT codigo FROM codigos_alternativos WHERE alternativo=?', (barcode.strip(),)).fetchone()
            if prior and prior[0] != code:
                raise ValueError(f'El código de barras ya pertenece al artículo {prior[0]}.')
        stamp = datetime.now().isoformat(timespec='seconds')
        with self.conn:
            self.conn.execute("INSERT INTO catalogo(codigo,descripcion,precio,en_tienda,actualizado_en) "
                "VALUES(?,?,?,?,?) ON CONFLICT(codigo) DO UPDATE SET "
                "descripcion=excluded.descripcion,precio=excluded.precio,actualizado_en=excluded.actualizado_en",
                (code, name.strip(), str(price), int(bool(in_store)),stamp))
            self.conn.execute('INSERT INTO precios_manuales(codigo,precio) VALUES (?,?) ON CONFLICT(codigo) DO UPDATE SET precio=excluded.precio', (code,str(price)))
            if barcode and barcode.strip() != code:
                self.conn.execute('INSERT INTO codigos_alternativos(alternativo,codigo) VALUES (?,?) ON CONFLICT(alternativo) DO UPDATE SET codigo=excluded.codigo',(barcode.strip(),code))
        if in_store:
            self.set_store([code], True, manual=True)

    def get_manual_price(self, code):
        r=self.conn.execute('SELECT precio FROM precios_manuales WHERE codigo=?',(code,)).fetchone()
        return money(r[0]) if r else None

    def get_product(self, code):
        r=self.conn.execute('SELECT codigo,descripcion,precio,en_tienda FROM catalogo WHERE codigo=?',(code,)).fetchone()
        return dict(code=r[0],name=r[1],price=money(r[2] or '0'),store=bool(r[3])) if r else None

    def backup(self, destination):
        with sqlite3.connect(str(destination)) as dest:
            self.conn.backup(dest)



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('1280x800')
        self.minsize(980, 620)
        self.configure(bg='#f3f5f9')
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.drive_folder = drive_folder_saved()
        self.search = tk.StringVar()
        self.kind = tk.StringVar(value='Precios modificados')
        self.changes = []
        self.new_items = []
        self.visible = {}
        self.chosen = set()
        self.current_lot = None
        self.qty = {}
        self.db = CatalogDB()
        # Import the checked 3,049 barcode↔IdArticulo matches automatically,
        # including on machines that update only comparador_cotyland.py via GitHub.
        self.barcode_seed_result = self.db.seed_verified_barcodes(verified_seed_records())
        self.in_store = self.db.store_codes()
        self.only_store = tk.BooleanVar(value=True)
        self._build()
        if self.barcode_seed_result['conflicts']:
            self.after(700, self.warn_barcode_conflicts)

    def warn_barcode_conflicts(self):
        conflicts = self.barcode_seed_result['conflicts']
        sample = '\n'.join(f'Barras {b}: archivo = {new}, base = {old}' for b,new,old in conflicts[:8])
        messagebox.showwarning('Asociaciones de barras para revisar',
           f'Se detectaron {len(conflicts)} asociaciones que contradicen las existentes.\n'
           'NO se modificaron las relaciones que tenías guardadas.\n\n' + sample +
           '\n\nAntes de imprimir esos artículos, verificá sus códigos en A-Sistemas.')

    def _build(self):
        style = ttk.Style(self)
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        style.configure('Treeview', rowheight=26, font=('Segoe UI', 9), background='white', fieldbackground='white')
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))
        style.map('Treeview', background=[('selected', '#ffdf69')], foreground=[('selected', '#161616')])
        style.configure('Compact.TButton', padding=(8, 4), font=('Segoe UI', 9))
        style.configure('Icon.TButton', padding=(6, 4), font=('Segoe UI Symbol', 9))
        style.configure('Primary.TButton', padding=(10, 5), font=('Segoe UI', 9, 'bold'))

        header = tk.Frame(self, bg='#ffe03a', padx=20, pady=10)
        header.pack(fill='x')
        ttk.Button(header, text='BUSCAR ACTUALIZACIONES', command=self.check_updates, style='Compact.TButton').pack(side='right')
        tk.Label(header, text='COTYLAND', font=('Segoe UI', 20, 'bold'), bg='#ffe03a').pack(anchor='w')
        tk.Label(header, text='Comparador de precios y etiquetas · 7 × 3,5 cm', font=('Segoe UI', 10), bg='#ffe03a').pack(anchor='w')

        files = tk.Frame(self, bg='#f3f5f9', padx=18, pady=10)
        files.pack(fill='x')
        self._file_row(files, 'LISTADO ANTERIOR', self.old_path, 0)
        self._file_row(files, 'LISTADO ACTUALIZADO', self.new_path, 1)
        files.columnconfigure(1, weight=1)
        actions = tk.Frame(files, bg='#f3f5f9')
        actions.grid(row=2, column=0, columnspan=3, sticky='ew', pady=(8, 0))
        ttk.Button(actions, text='COMPARAR LISTADOS', command=self.compare, style='Compact.TButton').pack(side='left', fill='x', expand=True)
        ttk.Button(actions, text='☁ CARGAR DESDE DRIVE', command=self.load_from_drive, style='Compact.TButton').pack(side='left', padx=(8, 0))
        ttk.Button(actions, text='📁', command=self.choose_drive_folder, style='Icon.TButton').pack(side='left', padx=(6, 0))

        bar = tk.Frame(self, bg='#f3f5f9', padx=18, pady=2)
        bar.pack(fill='x')
        tk.Label(bar, text='Ver:', bg='#f3f5f9').pack(side='left')
        cb = ttk.Combobox(bar, textvariable=self.kind, values=['Precios modificados', 'Artículos nuevos', 'Catálogo en tienda'], state='readonly', width=20)
        cb.pack(side='left', padx=(6, 10))
        cb.bind('<<ComboboxSelected>>', lambda _: self.refresh())
        ttk.Button(bar, text='Catálogo inicial', command=self.import_store_csv, style='Compact.TButton').pack(side='left', padx=(0, 10))
        self.new_count_label = tk.Label(bar, text='PRODUCTOS NUEVOS: 0', bg='#fff1b8', fg='#332700', font=('Segoe UI', 9, 'bold'), padx=8, pady=3, cursor='hand2')
        self.new_count_label.pack(side='left', padx=(0, 12))
        self.new_count_label.bind('<Button-1>', lambda _: self.show_new_products())
        tk.Label(bar, text='Buscar:', bg='#f3f5f9').pack(side='left')
        entry = ttk.Entry(bar, textvariable=self.search)
        entry.pack(side='left', fill='x', expand=True, padx=(6, 8))
        ttk.Checkbutton(bar, text='Solo productos en tienda', variable=self.only_store, command=self.refresh).pack(side='left')
        self.search.trace_add('write', lambda *_: self.refresh())

        seed = self.barcode_seed_result
        self.status = tk.StringVar(value=(f'Escáner: {seed["added"] + seed["already"]:,} códigos vinculados desde Libro1 '
              f'({seed["added"]:,} nuevos; {len(seed["conflicts"])} conflictos). '
              'Cargá los listados para actualizar precios.').replace(',', '.'))
        tk.Label(self, textvariable=self.status, bg='#f3f5f9', fg='#364152', font=('Segoe UI', 9, 'bold'), anchor='w', padx=18, pady=5).pack(fill='x')

        tools = tk.Frame(self, bg='#f3f5f9', padx=18, pady=2)
        tools.pack(fill='x')
        left_tools = tk.Frame(tools, bg='#f3f5f9')
        left_tools.pack(side='left', fill='x', expand=True)
        ttk.Button(left_tools, text='Marcar visibles', command=self.mark_visible, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Desmarcar', command=self.unmark, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Marcar tienda', command=lambda: self.set_store(True), style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Quitar tienda', command=lambda: self.set_store(False), style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Agregar nuevos', command=lambda: self.set_store(True), style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Importar lista', command=self.import_store_csv, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Copias…', command=self.set_copies, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='🔎 Buscar / escanear', command=self.scan_product, style='Compact.TButton').pack(side='left', padx=(0, 6))
        ttk.Button(left_tools, text='Importar barras', command=self.import_barcode_csv, style='Compact.TButton').pack(side='left', padx=(0, 6))

        right_tools = tk.Frame(tools, bg='#f3f5f9')
        right_tools.pack(side='right')
        ttk.Button(right_tools, text='🗄 Respaldo', command=self.backup_catalog, style='Compact.TButton').pack(side='right', padx=(6, 0))
        ttk.Button(right_tools, text='▌║█║ Etiquetas', command=self.labels, style='Primary.TButton').pack(side='right', padx=(6, 0))
        ttk.Button(right_tools, text='⤓ Exportar', command=self.export, style='Compact.TButton').pack(side='right', padx=(6, 0))

        frame = tk.Frame(self, padx=18, pady=8, bg='#f3f5f9')
        frame.pack(fill='both', expand=True)
        columns = ('marked', 'store', 'code', 'name', 'old', 'new', 'qty')
        self.table = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
        specs = [
            ('marked', 'Imprimir', 72, 'center'),
            ('store', 'En tienda', 72, 'center'),
            ('code', 'Código', 105, 'w'),
            ('name', 'Descripción', 560, 'w'),
            ('old', 'Precio anterior', 120, 'e'),
            ('new', 'Precio nuevo', 120, 'e'),
            ('qty', 'Copias', 60, 'center'),
        ]
        for key, title, width, anchor in specs:
            self.table.heading(key, text=title)
            self.table.column(key, width=width, minwidth=52, anchor=anchor, stretch=(key == 'name'))
        sc_y = ttk.Scrollbar(frame, orient='vertical', command=self.table.yview)
        sc_x = ttk.Scrollbar(frame, orient='horizontal', command=self.table.xview)
        self.table.configure(yscrollcommand=sc_y.set, xscrollcommand=sc_x.set)
        self.table.grid(row=0, column=0, sticky='nsew')
        sc_y.grid(row=0, column=1, sticky='ns')
        sc_x.grid(row=1, column=0, sticky='ew')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.table.bind('<Button-1>', self.click_checkbox)
        self.table.bind('<Double-1>', self.toggle_selected)
    def check_updates(self):
        import updater
        from tkinter.simpledialog import askstring
        url = updater.get_update_url()
        if not url:
            url = askstring('Configurar actualizaciones',
                'Para descargar futuras versiones automáticamente, necesitás una dirección HTTPS de actualizaciones.\n'
                'Pegá aquí la URL del archivo version.json publicado en tu servidor o GitHub:', parent=self)
            if not url: return
            try: updater.save_update_url(url)
            except ValueError as exc:
                messagebox.showerror('URL inválida', str(exc)); return
        try:
            result = updater.check_and_install(url)
            if result['status'] == 'current':
                messagebox.showinfo('Actualizado', 'Ya tenés la última versión disponible (' + APP_VERSION + ').')
            else:
                if messagebox.askyesno('Actualización instalada',
                        'Se instaló la versión ' + result['version'] + '.\n\n¿Reiniciar el programa ahora?'):
                    self.destroy()
                    subprocess.Popen([str(Path(sys.executable).with_name('pythonw.exe')) if sys.platform == 'win32' and Path(sys.executable).with_name('pythonw.exe').exists() else sys.executable, str(Path(__file__).resolve())], cwd=str(Path(__file__).resolve().parent))
        except Exception as exc:
            messagebox.showerror('No se pudo actualizar', str(exc) +
                '\n\nPodés cambiar la dirección editando actualizaciones.json junto al programa.')

    def _file_row(self, parent, label, var, row):
        tk.Label(parent, text=label, bg='#f3f5f9', font=('Segoe UI', 10, 'bold')).grid(row=row, column=0, padx=(0,10), pady=4, sticky='w')
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky='ew', pady=4)
        ttk.Button(parent, text='Examinar', command=lambda: self.browse(var), style='Compact.TButton').grid(row=row, column=2, padx=(10,0), pady=4)

    def browse(self, var):
        path = filedialog.askopenfilename(title='Seleccionar listado CSV', filetypes=[('CSV', '*.csv'), ('Todos los archivos','*.*')])
        if path: var.set(path)

    def choose_drive_folder(self):
        folder = filedialog.askdirectory(title='Elegí la carpeta de Google Drive que contiene viejo.csv y nuevo.csv',
                                        initialdir=str(self.drive_folder) if self.drive_folder.is_dir() else None)
        if folder:
            self.drive_folder = Path(folder)
            save_drive_folder(self.drive_folder)
            messagebox.showinfo('Carpeta de Drive guardada', str(self.drive_folder))
        return bool(folder)

    def load_from_drive(self):
        try:
            if not self.drive_folder.is_dir():
                if not self.choose_drive_folder(): return
            old, new = find_drive_pair(self.drive_folder)
            self.old_path.set(str(old))
            self.new_path.set(str(new))
            self.compare()
        except Exception as exc:
            messagebox.showerror('No se pudieron cargar las listas de Drive', str(exc) +
                '\n\nComprobá que Drive haya terminado de sincronizar y que ambos archivos sean del mismo lote.')

    @staticmethod
    def _list_pair_id(old_path, new_path):
        dig = hashlib.sha256()
        for label, filename in [('viejo', old_path), ('nuevo', new_path)]:
            dig.update(label.encode('ascii'))
            with open(filename, 'rb') as src:
                for chunk in iter(lambda: src.read(1024 * 1024), b''):
                    dig.update(chunk)
        return dig.hexdigest()

    def _persist_print(self, codes, enabled):
        # Save both for backwards compatibility. The important record is by
        # article + final price, not CSV bytes: re-exporting identical prices
        # must NOT recheck the label the user rejected.
        self.db.set_print_decisions(self.current_lot, codes, enabled)
        catalog = {row['new']['code']: row['new'] for row in self.changes + self.new_items}
        updates = []
        for code in codes:
            product = catalog.get(code) or self.db.get_product(code)
            if product:
                updates.append((code, product['price'], enabled))
        if updates:
            self.db.set_print_preferences(updates)

    def compare(self):
        if not self.old_path.get() or not self.new_path.get():
            messagebox.showwarning('Faltan archivos', 'Seleccioná los dos listados.'); return
        try:
            self.configure(cursor='watch'); self.update_idletasks()
            before = load_csv(self.old_path.get())
            after = load_csv(self.new_path.get())
            lot = self._list_pair_id(self.old_path.get(), self.new_path.get())
            self.changes = [dict(new=entry, old=before[code]['price']) for code, entry in after.items() if code in before and entry['price'] != before[code]['price']]
            self.new_items = [dict(new=entry, old=None) for code, entry in after.items() if code not in before]
            self.db.sync_catalog(after)
            self.in_store = self.db.store_codes()
            self.changes.sort(key=lambda item: item['new']['name'].casefold())
            self.new_items.sort(key=lambda item: item['new']['name'].casefold())
            self.new_count_label.config(text=f'PRODUCTOS NUEVOS: {len(self.new_items):,}'.replace(',', '.'))
            # Cada comparación nueva selecciona por defecto para IMPRIMIR
            # únicamente los PRECIOS MODIFICADOS de productos marcados EN TIENDA.
            # Las dos selecciones son independientes: destildar impresión no
            # quita el artículo de tienda, ni se vuelve a tildar al refrescar.
            self.current_lot = lot
            self.chosen = {item['new']['code'] for item in self.changes
                           if item['new']['code'] in self.in_store}
            # Previous versions saved per exact file pair; preserve their
            # decisions for that pair and migrate them to the stronger rule.
            legacy = self.db.print_decisions(lot)
            for code, enabled in legacy.items():
                if enabled:
                    self.chosen.add(code)
                else:
                    self.chosen.discard(code)
            valid = {row['new']['code']: row['new']['price'] for row in self.changes + self.new_items}
            if legacy:
                self.db.set_print_preferences((code,valid[code], enabled)
                                              for code,enabled in legacy.items() if code in valid)
            # Strongest rule: decisions persist for a given product / price
            # even after re-exporting either CSV, using another PC's copies,
            # or re-running the comparison.
            for code, enabled in self.db.print_preferences(self.changes + self.new_items).items():
                if enabled:
                    self.chosen.add(code)
                else:
                    self.chosen.discard(code)
            self.qty.clear()
            self.kind.set('Precios modificados'); self.search.set('')
            self.summary = f'{len(before):,} anteriores · {len(after):,} actuales · {len(self.changes):,} precios cambiados · {len(self.new_items):,} productos nuevos'.replace(',', '.')
            self.refresh()
        except Exception as e:
            messagebox.showerror('Error al leer los archivos', str(e))
        finally:
            self.configure(cursor='')

    def current(self):
        if self.kind.get() == 'Precios modificados': return self.changes
        if self.kind.get() == 'Artículos nuevos': return self.new_items
        return [dict(new=p,old=None) for p in self.db.products()]

    def refresh(self):
        for item in self.table.get_children(): self.table.delete(item)
        q = self.search.get().strip().casefold()
        self.visible = {}
        for item in self.current():
            entry = item['new']; code = entry['code']
            if q and q not in code.casefold() and q not in entry['name'].casefold(): continue
            if self.kind.get() == 'Precios modificados' and self.only_store.get() and code not in self.in_store: continue
            # New items always visible so user can decide which to add.
            self.visible[code] = item
            self.table.insert('', 'end', iid=code, values=('☑' if code in self.chosen else '□', '☑' if code in self.in_store else '□', code, entry['name'], display_money(item['old']) if item['old'] is not None else '—', display_money(entry['price']), self.qty.get(code,1)))
        self.status.set((getattr(self,'summary','Sin comparación') + f'   |   Mostrando {len(self.visible)}   |   En tienda {len(self.in_store)}   |   Para imprimir {len(self.chosen)}'))

    def save_store(self, codes, value):
        try:
            self.db.set_store(codes, value)
            self.in_store = self.db.store_codes()
        except (OSError, sqlite3.Error) as exc:
            messagebox.showerror('No se pudo guardar la base de datos', str(exc))

    def backup_catalog(self):
        path = filedialog.asksaveasfilename(title='Guardar copia de seguridad del catálogo',
            defaultextension='.sqlite3', initialfile='Cotyland_catalogo_respaldo.sqlite3',
            filetypes=[('Base de datos SQLite','*.sqlite3')])
        if not path: return
        try:
            self.db.backup(path)
            messagebox.showinfo('Respaldo guardado', f'Se guardó una copia del catálogo en:\n{path}')
        except (OSError,sqlite3.Error) as exc:
            messagebox.showerror('Error de respaldo', str(exc))

    def show_new_products(self):
        self.kind.set('Artículos nuevos')
        self.refresh()

    def import_store_csv(self):
        """Import store codes without clearing the old selection. Accepts Cotyland's
        repeated-header 16-column CSV, simple CSV with Codigo column, or a stock CSV.
        An active flag or a price is never automatically interpreted as stock.
        """
        path = filedialog.askopenfilename(title='Elegir catálogo inicial CSV o listado con stock',
            filetypes=[('Archivos CSV','*.csv'),('Todos los archivos','*.*')])
        if not path: return
        try:
            self.configure(cursor='watch'); self.update_idletasks()
            rows, origin, stock_present = self._read_store_import(path)
            if not rows: raise ValueError('No se encontraron códigos válidos.')
        except (OSError, ValueError, csv.Error) as exc:
            messagebox.showerror('No se pudo leer el listado', str(exc)); return
        finally:
            self.configure(cursor='')

        # For stock files, any non-zero quantity (including negative) qualifies.
        # For an active-only export, require an explicit opt-in to provisional marking.
        with_stock = sum(1 for p in rows.values() if p.get('stock') is not None and p['stock'] != 0)
        priced = sum(1 for p in rows.values() if p.get('price') is not None and p['price'] > 0)
        zero_stock = sum(1 for p in rows.values() if p.get('stock') == 0)
        dialog = tk.Toplevel(self)
        dialog.title('Confirmar importación · Cotyland')
        dialog.geometry('720x455')
        dialog.minsize(640,400)
        dialog.transient(self); dialog.grab_set()
        container = ttk.Frame(dialog, padding=20); container.pack(fill='both',expand=True)
        ttk.Label(container,text=f'Archivo: {Path(path).name}',font=('Segoe UI',12,'bold')).pack(anchor='w',pady=(0,8))
        ttk.Label(container,text=f'Códigos encontrados: {len(rows):,} | Con precio positivo: {priced:,}' +
                  (f' | Stock distinto de cero: {with_stock:,} | Stock cero: {zero_stock:,}' if stock_present else ''),
                  font=('Segoe UI',10)).pack(anchor='w',pady=(0,9))
        note = ('Se detectó una columna de existencias. Se pueden marcar automáticamente los '
                'productos con stock positivo o negativo; los de stock cero NO se desmarcan.') if stock_present else (
                'ATENCIÓN: este archivo NO tiene una columna verificable de stock. '
                'ACTIVO o PRECIO > 0 NO significa necesariamente que el producto esté físicamente en tienda. '
                'Podés usarlo como selección preliminar, pero conviene revisar o conseguir el reporte de existencias.')
        ttk.Label(container,text=note,wraplength=665,justify='left',foreground='#a53b17').pack(anchor='w',pady=(0,12))
        mode = tk.StringVar(value='stock' if stock_present else 'none')
        if stock_present:
            ttk.Radiobutton(container,variable=mode,value='stock',text=f'Marcar solo stock distinto de 0 (incluye negativos): {with_stock:,}').pack(anchor='w',pady=3)
        ttk.Radiobutton(container,variable=mode,value='none',text='Solo cargar datos de artículos, SIN marcar ninguno en tienda').pack(anchor='w',pady=3)
        ttk.Radiobutton(container,variable=mode,value='priced',text=f'Marcar como selección PROVISIONAL los {priced:,} artículos con precio mayor a cero').pack(anchor='w',pady=3)
        ttk.Radiobutton(container,variable=mode,value='all',text=f'Marcar TODOS los {len(rows):,} códigos del archivo (solo si realmente son de la tienda)').pack(anchor='w',pady=3)
        ttk.Label(container,text='Esta carga es PROVISIONAL. Las marcas existentes se conservan; las decisiones manuales tienen prioridad.',wraplength=665).pack(anchor='w',pady=(12,4))
        def apply():
            selection = mode.get()
            codes = ([p['code'] for p in rows.values() if p.get('stock') is not None and p['stock'] != 0]
                     if selection == 'stock' else
                     [p['code'] for p in rows.values() if p.get('price') is not None and p['price'] > 0]
                     if selection == 'priced' else list(rows) if selection == 'all' else [])
            if selection in ('priced','all'):
                if not messagebox.askyesno('Confirmación especial',
                    f'Vas a marcar {len(codes):,} artículos como «En tienda» SIN verificar el stock.\n\n'
                    '¿Confirmás esta selección? Se sumará a lo ya marcado.', parent=dialog):
                    return
            elif selection == 'stock':
                if not messagebox.askyesno('Confirmar stock',
                    f'Se van a marcar {len(codes):,} artículos con stock distinto de cero.\n'
                    'No se quitará ninguna marca existente. ¿Continuar?', parent=dialog):
                    return
            try:
                self.db.sync_catalog({p['code']:dict(code=p['code'],name=p['name'],price=p.get('price') or Decimal(0)) for p in rows.values()})
                previous_store = set(self.in_store)
                if codes: self.db.set_store(codes, True, manual=False)
                self.in_store = self.db.store_codes()
                newly_in_store = self.in_store - previous_store
                already_rejected = {code for code, enabled in self.db.print_preferences(self.changes).items() if not enabled}
                already_rejected.update(code for code, enabled in self.db.print_decisions(self.current_lot).items() if not enabled)
                self.chosen.update(item['new']['code'] for item in self.changes
                                   if item['new']['code'] in newly_in_store and item['new']['code'] not in already_rejected)
                dialog.destroy()
                self.refresh()
                messagebox.showinfo('Importación realizada',
                    f'Artículos importados a la base: {len(rows):,}\n'
                    f'Códigos marcados en esta operación: {len(codes):,}\n'
                    f'Total guardado «En tienda»: {len(self.in_store):,}\n\n'
                    'Las marcas seguirán guardadas al actualizar el comparador.')
            except (OSError,sqlite3.Error) as exc:
                messagebox.showerror('Error al guardar la base',str(exc),parent=dialog)
        buttons=ttk.Frame(container); buttons.pack(anchor='e',side='bottom',pady=(10,0))
        ttk.Button(buttons,text='Cancelar',command=dialog.destroy).pack(side='right',padx=6)
        ttk.Button(buttons,text='IMPORTAR',command=apply).pack(side='right')

    @staticmethod
    def _read_store_import(path):
        """Returns dictionary code -> {code,name,price,stock} and whether stock exists."""
        result={}; stock_found=False
        with open(path,'r',encoding='utf-8-sig',errors='replace',newline='') as f:
            sample=f.readline(); f.seek(0)
            reader=csv.reader(f,delimiter=';' if sample.count(';') > sample.count(',') else ',')
            first=next(reader, None)
            if first is None: return {},'empty',False
            # Standard export: repeated label columns + real values at 9,10,14.
            cotyland=(len(first)>=15 and first[4].strip().lower()=='idarticulo'
                      and first[5].strip().lower()=='descripcion')
            if cotyland:
                def parse_row(row):
                    if len(row)<15 or row[4].strip().lower()!='idarticulo': return
                    code=row[9].strip()
                    if not code:return
                    result[code]={'code':code,'name':row[10].strip(), 'price':money(row[14]),'stock':None}
                parse_row(first)
                for row in reader:parse_row(row)
                return result,'Cotyland',False
            names=[c.strip().casefold().replace(' ','').replace('_','') for c in first]
            code_names={'idarticulo','codigo','código','sku','codarticulo','codigodearticulo'}
            code_idx=next((i for i,c in enumerate(names) if c in code_names),None)
            if code_idx is None: raise ValueError('No se identificó la columna de código. Usá IdArticulo, Codigo o SKU.')
            stock_names={'stock','existencias','existencia','stockactual','cantidadstock','cantidadenstock','saldo','stockdisponible'}
            stock_idx=next((i for i,c in enumerate(names) if c in stock_names),None)
            stock_found=stock_idx is not None
            price_idx=next((i for i,c in enumerate(names) if c in ('precio','precioventa','preciolista')),None)
            name_idx=next((i for i,c in enumerate(names) if c in ('descripcion','descripción','nombre','producto')),None)
            for line,row in enumerate(reader,2):
                if len(row)<=code_idx:continue
                code=row[code_idx].strip()
                if not code:continue
                p=money(row[price_idx]) if price_idx is not None and len(row)>price_idx else None
                qty=None
                if stock_found and len(row)>stock_idx:
                    raw=row[stock_idx].strip().replace(' ','')
                    if raw:
                        try:
                            if ',' in raw:raw=raw.replace('.','').replace(',','.')
                            qty=Decimal(raw)
                        except InvalidOperation:
                            raise ValueError(f'Stock inválido para el artículo {code} en fila {line}: {row[stock_idx]}')
                result[code]={'code':code,'name':row[name_idx].strip() if name_idx is not None and len(row)>name_idx else '',
                              'price':p,'stock':qty}
        return result,'simple',stock_found

    def import_barcode_csv(self):
        path = filedialog.askopenfilename(title='Importar códigos de barras e IdArticulo',
                filetypes=[('Listados CSV', '*.csv'), ('Todos los archivos', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8-sig', newline='') as f:
                sample = f.readline()
                f.seek(0)
                reader = csv.DictReader(f, delimiter=';' if sample.count(';') > sample.count(',') else ',')
                names = {str(k).strip().lower().replace(' ', '').replace('_', ''): k for k in reader.fieldnames or []}
                sku = next((names[n] for n in ('idarticulo', 'sku', 'codigointerno', 'codigo') if n in names), None)
                bar = next((names[n] for n in ('codigobarra', 'codigobarras', 'códigobarra', 'códigobarras', 'barcode', 'ean', 'ean13') if n in names), None)
                if not sku or not bar:
                    raise ValueError('El CSV debe tener las columnas IdArticulo y Codigo_Barra (o similares).')
                pairs = [(row.get(sku, ''), row.get(bar, '')) for row in reader]
            if not pairs:
                raise ValueError('El archivo no tiene registros.')
            if not messagebox.askyesno('Importar códigos de barras',
                    f'Se encontraron {len(pairs):,} filas.\n\nSe agregarán asociaciones código de barras ↔ código interno, sin reemplazar las existentes que correspondan a otro artículo. ¿Continuar?'):
                return
            count = self.db.import_barcodes(pairs)
            messagebox.showinfo('Códigos de barras importados',
                    f'Se incorporaron {count:,} asociaciones de código de barras. Ya podés escanear esos artículos.')
        except (OSError, ValueError, csv.Error, sqlite3.Error) as exc:
            messagebox.showerror('No se pudo importar el archivo', str(exc))

    def scan_product(self):
        """Capture one or many barcode scans. Never resolve scanned values by fuzzy SKU search."""
        dlg = tk.Toplevel(self)
        dlg.title('Cotyland · Escaneo por código de barras')
        dlg.geometry('1020x660')
        dlg.minsize(810, 510)
        dlg.transient(self)
        panel = ttk.Frame(dlg, padding=14)
        panel.pack(fill='both', expand=True)
        ttk.Label(panel, text='ESCANEO DE ARTÍCULOS — UNO O VARIOS',
                  font=('Segoe UI', 13, 'bold')).pack(anchor='w')
        ttk.Label(panel, text='Escaneá productos uno tras otro: cada lectura se reconoce por código de barras REAL y el campo se limpia solo. El lector debe enviar el código completo. Para escribir manualmente, Enter.',
                  wraplength=920).pack(anchor='w', pady=(4, 8))
        scan_row = ttk.Frame(panel)
        scan_row.pack(fill='x')
        ttk.Label(scan_row, text='Código de barras:').pack(side='left')
        scan_value = tk.StringVar()
        scan_entry = ttk.Entry(scan_row, textvariable=scan_value, font=('Segoe UI', 13))
        scan_entry.pack(side='left', fill='x', expand=True, padx=9)
        auto_capture = tk.BooleanVar(value=True)
        ttk.Checkbutton(panel, text='Lectura continua (procesar automáticamente al terminar el escaneo)',
                        variable=auto_capture).pack(anchor='w', pady=(5, 1))
        ttk.Label(panel, text='Si el lector agrega Enter, se procesa inmediatamente. Si no, se procesa luego de una pausa de 450 ms. Desactivá la lectura continua para escribir despacio a mano.',
                  wraplength=920).pack(anchor='w', pady=(0, 5))
        auto_print = tk.BooleanVar(value=True)
        ttk.Checkbutton(panel, text='Preparar automáticamente una etiqueta para cada artículo reconocido',
                        variable=auto_print).pack(anchor='w', pady=(6, 8))

        cols = ('barcode', 'code', 'name', 'price', 'state', 'count')
        tree_box = ttk.Frame(panel)
        tree_box.pack(fill='both', expand=True)
        tree = ttk.Treeview(tree_box, columns=cols, show='headings', selectmode='browse')
        for key, label, width in [('barcode','Barras leído',145),('code','IdArticulo',110),
                                  ('name','Artículo',380),('price','Precio',115),
                                  ('state','Estado',125),('count','Lecturas',70)]:
            tree.heading(key, text=label)
            tree.column(key, width=width, anchor='e' if key=='price' else 'w', stretch=key=='name')
        tree.column('count', anchor='center')
        scroll = ttk.Scrollbar(tree_box, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # Keys are the exact barcode strings read, not generated SKU approximations.
        session = {}
        scanned_copies = {}
        info = tk.StringVar(value='0 lecturas. Código desconocido = pendiente, sin asignación automática.')
        ttk.Label(panel, textvariable=info, wraplength=960).pack(anchor='w', pady=8)

        def display():
            selected = tree.selection()
            tree.delete(*tree.get_children())
            for barcode, item in session.items():
                product = item.get('product')
                status = 'Verificado' if product else 'PENDIENTE'
                tree.insert('', 'end', iid=barcode, values=(barcode,
                    product['code'] if product else '—',
                    product['name'] if product else 'SIN IDENTIFICAR',
                    display_money(product['price']) if product else '—',
                    status, item['count']))
            for selection in selected:
                if tree.exists(selection):
                    tree.selection_set(selection)
            pending = sum(1 for x in session.values() if not x['product'])
            info.set(f'{len(session)} códigos diferentes · {sum(x["count"] for x in session.values())} lecturas · '
                     f'{pending} sin verificar. Los desconocidos NO se agregan ni imprimen.')

        def capture(_event=None):
            barcode = scan_value.get().strip()
            scan_value.set('')
            if not barcode:
                scan_entry.focus_set()
                return
            product = self.db.resolve_verified_barcode(barcode)
            if barcode in session:
                session[barcode]['count'] += 1
                if product:
                    session[barcode]['product'] = product
            else:
                session[barcode] = {'count': 1, 'product': product}
            if product and auto_print.get():
                code = product['code']
                self.chosen.add(code)
                # Count repeated scans of the same physical code as additional copies.
                scanned_copies[code] = scanned_copies.get(code, 0) + 1
                self.qty[code] = max(self.qty.get(code, 0), scanned_copies[code])
                self._persist_print([code], True)
                self.refresh()
            display()
            tree.selection_set(barcode)
            tree.see(barcode)
            scan_entry.focus_set()

        # Los lectores USB son teclados: unos agregan Enter y otros no.
        # Con lectura continua, capturar al detectar una pausa entre teclas.
        # El temporizador se cancela con cada tecla para no cortar el código.
        capture_job = None
        def cancel_capture_job():
            nonlocal capture_job
            if capture_job is not None:
                try:
                    dlg.after_cancel(capture_job)
                except tk.TclError:
                    pass
                capture_job = None

        def capture_now(_event=None):
            cancel_capture_job()
            capture()
            return 'break'  # Evita ejecutar dos veces la misma lectura.

        def schedule_capture(*_args):
            nonlocal capture_job
            cancel_capture_job()
            if auto_capture.get() and scan_value.get().strip():
                capture_job = dlg.after(450, capture_now)

        scan_value.trace_add('write', schedule_capture)
        scan_entry.bind('<Return>', capture_now)
        auto_capture.trace_add('write', schedule_capture)
        dlg.bind('<Destroy>', lambda ev: cancel_capture_job() if ev.widget == dlg else None)

        def print_scanned_only():
            codes = {code for code, count in scanned_copies.items() if count > 0}
            if not codes:
                messagebox.showwarning('Sin etiquetas escaneadas',
                    'Primero escaneá productos reconocidos con la opción de preparar etiqueta activa,\n'
                    'o seleccioná un producto y tocá PREPARAR ETIQUETA.', parent=dlg)
                scan_entry.focus_set()
                return
            # Scan output is independent of any other changes selected in the main window.
            self.labels(only_codes=codes)
            scan_entry.focus_set()

        def chosen_barcode():
            row = tree.selection()
            if not row:
                messagebox.showinfo('Elegí un código', 'Seleccioná una lectura en la tabla.', parent=dlg)
                return None
            return row[0]

        def link_selected():
            barcode = chosen_barcode()
            if not barcode:
                return
            from tkinter.simpledialog import askstring
            code = askstring('Vincular con A-Sistemas',
                    'Código de barras leído: ' + barcode + '\n\nIngresá el IdArticulo EXACTO que muestra A-Sistemas:',
                    parent=dlg)
            if not code:
                return
            code = code.strip()
            product = self.db.get_product(code)
            if not product:
                messagebox.showerror('Artículo no encontrado',
                    'Ese IdArticulo no existe en el catálogo. No se guardó ninguna relación.', parent=dlg)
                return
            if not messagebox.askyesno('Confirmar equivalencia real',
                    'Código de barras: ' + barcode + '\nCódigo interno: ' + code +
                    '\nDescripción: ' + product['name'] +
                    '\n\n¿Verificaste en A-Sistemas que este código de barras corresponde exactamente a ese artículo?',
                    parent=dlg):
                return
            try:
                self.db.link_barcode(barcode, code)
                session[barcode]['product'] = self.db.resolve_verified_barcode(barcode)
                display()
                tree.selection_set(barcode)
                info.set('Equivalencia verificada y guardada. Podés escanear el siguiente producto.')
            except (ValueError, sqlite3.Error) as exc:
                messagebox.showerror('No se pudo vincular', str(exc), parent=dlg)
            scan_entry.focus_set()

        def prepare_selected():
            barcode = chosen_barcode()
            if not barcode:
                return
            product = session[barcode]['product']
            if not product:
                messagebox.showwarning('Código no verificado',
                    'No se puede preparar una etiqueta sin confirmar la relación código de barras → IdArticulo.', parent=dlg)
                return
            code = product['code']
            self.chosen.add(code)
            scanned_copies[code] = scanned_copies.get(code, 0) + 1
            self.qty[code] = max(self.qty.get(code, 0), scanned_copies[code])
            self._persist_print([code], True)
            self.refresh()
            info.set('Etiqueta preparada para ' + product['name'] + '.')
            scan_entry.focus_set()

        def sku_search():
            """Explicit separate SKU/manual lookup: never used to interpret a barcode scan."""
            from tkinter.simpledialog import askstring
            term = askstring('Buscar por código interno o descripción',
                             'Ingresá un IdArticulo/SKU o una descripción (NO código de barras):', parent=dlg)
            if not term:
                return
            matches = self.db.find_products(term, limit=15)
            if not matches:
                messagebox.showinfo('Sin resultados', 'No hay coincidencias en el catálogo.', parent=dlg)
                return
            msg = '\n'.join(f'{r[0]} — {r[1]}' for r in matches[:10])
            messagebox.showinfo('Artículos encontrados', msg +
                '\n\nPara vincular barras, verificá primero el IdArticulo en A-Sistemas y usá VINCULAR.', parent=dlg)
            scan_entry.focus_set()

        def add_manual():
            """Manual product creation is explicit; no SKU is inferred from barcode."""
            from tkinter.simpledialog import askstring
            code = askstring('Nuevo artículo', 'Ingresá el IdArticulo REAL de A-Sistemas:', parent=dlg)
            if not code:
                return
            code = code.strip()
            name = askstring('Nuevo artículo', 'Descripción EXACTA del artículo:', parent=dlg)
            if not name:
                return
            price = askstring('Nuevo artículo', 'Precio de venta actual:', parent=dlg)
            if price is None:
                return
            try:
                value = money(price)
                if value < 0:
                    raise ValueError('El precio no puede ser negativo.')
                if not messagebox.askyesno('Confirmar artículo',
                    f'IdArticulo: {code}\nDescripción: {name}\nPrecio: {display_money(value)}\n\n'
                    '¿Confirmaste estos datos contra A-Sistemas?', parent=dlg):
                    return
                self.db.save_product(code, name, price, True)
                self.in_store = self.db.store_codes()
                self.refresh()
                info.set('Artículo ' + code + ' guardado con los datos ingresados. Para asociar barras, escaneá y vinculá.')
            except (ValueError, sqlite3.Error, OSError) as exc:
                messagebox.showerror('Error al guardar artículo', str(exc), parent=dlg)
            scan_entry.focus_set()

        buttons = ttk.Frame(panel)
        buttons.pack(fill='x', pady=(10, 0))
        ttk.Button(buttons, text='AGREGAR LECTURA', command=capture).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='VINCULAR BARRAS CON ARTÍCULO', command=link_selected).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='PREPARAR ETIQUETA', command=prepare_selected).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='Buscar por SKU (separado)', command=sku_search).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='AGREGAR ARTÍCULO', command=add_manual).pack(side='left', padx=(0, 7))
        ttk.Button(buttons, text='IMPRIMIR SOLO ESCANEADOS', command=print_scanned_only,
                   style='Primary.TButton').pack(side='right', padx=(8, 0))
        ttk.Button(buttons, text='CERRAR', command=dlg.destroy).pack(side='right')
        scan_entry.focus_set()

    def click_checkbox(self, event):
        row = self.table.identify_row(event.y)
        col = self.table.identify_column(event.x)
        if not row or self.table.identify_region(event.x, event.y) != 'cell': return
        if col == '#1':
            self.chosen.symmetric_difference_update({row})
            self._persist_print([row], row in self.chosen)
        elif col == '#2':
            adding = row not in self.in_store
            self.save_store([row], adding)
            if adding and row in {x['new']['code'] for x in self.changes + self.new_items}:
                self.chosen.add(row)
                self._persist_print([row], True)
        else: return
        self.refresh()

    def set_store(self, value):
        codes = set(self.table.selection())
        if not codes:
            messagebox.showinfo('Seleccioná filas', 'Seleccioná una o más filas con un clic (Ctrl o Shift para varias) y después usá este botón. También podés pulsar directamente la casilla En tienda.'); return
        self.save_store(codes, value)
        if value:
            newly_chosen = codes & {x['new']['code'] for x in self.changes + self.new_items}
            self.chosen.update(newly_chosen)
            self._persist_print(newly_chosen, True)
        self.refresh()

    def toggle_selected(self, event):
        row = self.table.identify_row(event.y)
        if not row or self.table.identify_column(event.x) in ('#1','#2'): return
        if row in self.chosen: self.chosen.remove(row)
        else: self.chosen.add(row)
        self._persist_print([row], row in self.chosen)
        self.refresh()

    def mark_visible(self):
        self.chosen.update(self.visible)
        self._persist_print(self.visible, True)
        self.refresh()

    def unmark(self):
        self._persist_print(self.chosen, False)
        self.chosen.clear(); self.refresh()

    def set_copies(self):
        from tkinter.simpledialog import askinteger
        number = askinteger('Cantidad de etiquetas', '¿Cuántas etiquetas por producto seleccionado?', initialvalue=1, minvalue=1, maxvalue=1000, parent=self)
        if number is not None:
            for code in self.chosen: self.qty[code] = number
            self.refresh()

    def export(self):
        if not self.changes and not self.new_items:
            messagebox.showinfo('Sin datos', 'Primero compará los listados.'); return
        path = filedialog.asksaveasfilename(defaultextension='.csv', initialfile='cotyland_cambios.csv', filetypes=[('CSV','*.csv')])
        if not path: return
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f, delimiter=';')
            w.writerow(['Tipo','IdArticulo','Descripcion','Precio anterior','Precio nuevo'])
            for type_name, rows in [('Cambio de precio',self.changes),('Artículo nuevo',self.new_items)]:
                for item in rows:
                    w.writerow([type_name,item['new']['code'],item['new']['name'],str(item['old']) if item['old'] is not None else '',str(item['new']['price'])])
        messagebox.showinfo('Exportación completa', f'Archivo guardado en:\n{path}')

    def labels(self, only_codes=None):
        target_codes = set(self.chosen) if only_codes is None else set(only_codes)
        if not target_codes:
            messagebox.showwarning('Seleccioná productos', 'Marcá los productos que quieras imprimir (doble clic o Marcar visibles).'); return
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import mm
            from reportlab.pdfbase.pdfmetrics import stringWidth
        except ImportError:
            messagebox.showerror('Falta dependencia', 'Instalá ReportLab con: py -m pip install reportlab'); return
        # Evitar impresiones accidentales de decenas de miles de páginas.
        rows = {x['new']['code']:x for x in self.changes + self.new_items}
        for code in target_codes:
            if code not in rows:
                p=self.db.get_product(code)
                if p:rows[code]={'new':p,'old':None}
        selected = []
        for code in target_codes:
            if code not in rows: continue
            item = dict(rows[code]); item['new'] = dict(item['new'])
            manual_price = self.db.get_manual_price(code)
            if manual_price is not None: item['new']['price'] = manual_price
            selected.append(item)
        without_price = [x['new']['code'] for x in selected if x['new']['price'] <= 0]
        if without_price:
            messagebox.showerror('Precios faltantes',
                'Estos artículos no tienen un precio vigente confirmado: '
                + ', '.join(without_price[:12]) +
                '.\n\nCompará los listados actuales antes de imprimir. No se genera un precio inventado.')
            return
        copies = sum(self.qty.get(x['new']['code'], 1) for x in selected)
        if copies > 5000 and not messagebox.askyesno('Muchas etiquetas',f'Se van a generar {copies} etiquetas. ¿Continuar?'): return
        path = filedialog.asksaveasfilename(defaultextension='.pdf', initialfile='etiquetas_cotyland_A4_24.pdf', filetypes=[('PDF','*.pdf')])
        if not path: return
        # A4 de 210 x 297 mm: 3 columnas x 8 filas de 70 x 35 mm.
        # Sin margen horizontal ni espacios entre etiquetas. Sobra una franja
        # vertical inevitable de 17 mm, que se reparte arriba y abajo.
        page_w, page_h = 210*mm, 297*mm
        w, h = 70*mm, 35*mm
        columns, per_page = 3, 24
        bottom = (page_h - 8*h)/2
        c = canvas.Canvas(path, pagesize=(page_w, page_h), pageCompression=1)
        generated_date = datetime.now().strftime('%d/%m/%Y')

        def clip(text, font, size, width):
            text = ' '.join(str(text).split())
            if stringWidth(text, font, size) <= width:
                return text
            while text and stringWidth(text+'…', font, size) > width:
                text = text[:-1]
            return text.rstrip() + '…'

        def description_lines(txt):
            # Hasta 2 líneas en el ancho útil. Evita cortar de golpe los nombres largos.
            words = ' '.join(str(txt or 'ARTÍCULO SIN DESCRIPCIÓN').split()).split()
            lines = ['']
            for word in words:
                try_line = (lines[-1]+' '+word).strip()
                if stringWidth(try_line, 'Helvetica-Bold', 8.2) <= w-5*mm:
                    lines[-1] = try_line
                elif len(lines) < 2:
                    lines.append(word)
                else:
                    lines[-1] = clip(lines[-1] + ' ' + word, 'Helvetica-Bold', 8.2, w-5*mm)
                    break
            return lines

        index = 0
        for item in sorted(selected, key=lambda x: x['new']['name'].casefold()):
            product = item['new']
            for _ in range(self.qty.get(product['code'], 1)):
                if index and index % per_page == 0:
                    c.showPage()
                slot = index % per_page
                col, row = slot % columns, slot // columns
                x = col*w
                y = page_h-bottom-(row+1)*h
                c.setLineWidth(0.27)
                c.setStrokeColorRGB(.65,.65,.65)
                c.rect(x,y,w,h,stroke=1,fill=0)
                c.setFillColorRGB(0,0,0)

                lines = description_lines(product['name'])
                c.setFont('Helvetica-Bold', 8.2)
                for n, line in enumerate(lines):
                    c.drawCentredString(x+w/2, y+h-(3.8+n*3.3)*mm, line)

                # Precio protagonista: sin barras, centrado y siempre dentro de la etiqueta.
                # El tamaño máximo real depende del ancho de los 70 mm.
                price = display_money(product['price'])
                max_price_width = w - 4*mm
                desired_size = 48.0
                size = min(desired_size, desired_size * max_price_width /
                           max(stringWidth(price, 'Helvetica-Bold', desired_size), 1))
                c.setFont('Helvetica-Bold', size)
                # Línea de precio entre la descripción y el código numérico.
                c.drawCentredString(x+w/2, y+12.0*mm, price)

                # Código exclusivamente numérico, sin barras ni fuente de barcode.
                identifier = str(product['code']).strip()
                c.setFont('Helvetica-Bold', 10.0)
                c.drawCentredString(x+w/2, y+5.1*mm,
                    clip('Cód. art.: ' + identifier, 'Helvetica-Bold', 10.0, w-4*mm))
                c.setFont('Helvetica', 7.0)
                c.drawCentredString(x+w/2, y+1.9*mm, 'Actualizado: ' + generated_date)
                index += 1
        c.save()
        if messagebox.askyesno('Etiquetas generadas',f'Guardadas {copies} etiquetas en hojas A4 (24 etiquetas por página, 70 × 35 mm cada una).\n\n¿Abrir el PDF para imprimir?'):
            try:
                if sys.platform == 'win32': os.startfile(path)
                elif sys.platform == 'darwin': subprocess.Popen(['open',path])
                else: subprocess.Popen(['xdg-open',path])
            except Exception as e: messagebox.showinfo('Archivo generado',f'Encontralo en: {path}\n{e}')

if __name__=='__main__':
    App().mainloop()

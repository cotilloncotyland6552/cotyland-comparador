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
APP_VERSION = '7.5.9'
# Asociaciones del Libro1.xlsx recuperadas de filas normales y texto tabulado incrustado.
# Verificadas contra SKU y descripcion de catalogos reales de Cotyland.
# NO deducir ceros iniciales, NO transformar ni inventar codigos.
_BARCODE_DATA_B85 = (
    'c-o}BTXWk;wl4Tr5c}bT6w#KDR{&2FAVo>ENrKIbTk<*cpmBMsJzXwW*sdKs`(NysKeFfbJa+%FS(nU}2@s?tpN?u5s|xEwA}=e~ef{_U{%@uHPpN<V'
    '%fG2o<L?T7F5Attz3Dc^y4|d2tKw!>+;?x=ZLxL#_J95*r>sh$)JS=^X;-twqAiMpdUw}<Ser&3_6~kdyUl90Ym3`Gkj#qhe7T0d-7fZB@&4)C=da(3'
    'k6(ZL`t8&A!j>Pu^ubJ--drK>dcWAtHZ6|zu3NUnv|AOE&3v)wU?^pu*m@I*;Sc=aZV|q5(ZYZ>vzu;p(-rMvw+Bd778(WsIDME(%Lge_!_RipZHpXe'
    '_0RvMY~hq25c))D{Ru+t3#aXXuyR7^6s^?it}S*mnDLb_k49K$BXrHF5!Numx<{x=8Ws>5UDHgrtKF=7+rjE|cXzYRmY1nG^uJ6gQ^gl*t?7kJ+DSmS'
    'nQh~S+V;n(_*o`1_+2;fMXO!Apk-XQ|N2kZA|3pgZ`&2DEqujpeo)1Qc9#PrQzj(Cn{0=JFP8hoZjKua_b5yO{MAlpn+{-{_Cpk#P+YdVvRwtE*=^=<'
    'REmQclDjkri$8DXt8UxwH}fu|QhGpDXH+;J>v^~8R`}YqT(&UwpzYu|D3caiTl_icmXmgR-{Oq7#a+9f-gjZ%3Rv*nU;%IgMEa_d6VWdA%lQgN-N6aP'
    '@4uf<+u=4--ltV>)QOS%!s>(3!;x1i?KDMObi3ce^V|Il7OJ?#ZLpoU*NgqC#T|4sZIza_G9|6meA@A6Z;-D584P>B!JV(%@iB&NBAd_9w~wAi4ZB1S'
    '2-AX@TFIortZx?GZkJ6TPFq)KeNbv-08J^=XKD8{V}z|nhb-*Z0ZTMyFp~fi*0A=5mh*A8rHPp1to`HH-qSEu!UHYsm?>-#+SoTUm~!BLZ?E1G;~kFI'
    'C?(%*9nE+PD|Wl-fD7Hu-?qbv8`5hhT+9+DoWKVX=U#8-tLbdp9+Naib_v}%@8Ln*-3xd%Zpk(9hus{u3ZL_72mGcu*gvoOI9ACyu{E*m#d5Z~*$_Yr'
    'gE$OsEz=ksR;O}?_W87fwcNCD%fTegw(Hq+4vcC73=95!TksVKtkDgitghsqb+yZ<oS0Oxqd!kgyY*Ja?v;7btz}1G>NpD6CN!0k7Gdu8b6kWCpcTMl'
    '7ZaEZH`p}JB=%mFcF0{H3)~2^#eCP{%GX7GsE-HW>RJwLWpNtbwi_7la&|L^n*s)YtqZ`|bh5fVdS9mFQqiS0n1^+{Di-Y&Z;x<vOnGQP2sj>#*IJf@'
    'FXi?<u&yb5!S*zN_fDlRpk1_`?Fzn%?`O+L*vT+2Eqx5#U$7mw!&g_<hF2Tl9tZ+oJ8k=)F=_AiGvp+jt7(Dv-S;0~KmMbC<yDoE<5V~Hn0Es=!voT8'
    'r~4Kr7I?xytf;M)1KrS;7!?!B!Ve~*u98UCH0Jo#ayp0|;}_lR7t<NsK*wuQS8+mY?VI`x7@iZjTdotH*2j2_iu<I7ljDg^(bnV}anVgWS{=9>*0WuZ'
    'J7JfU#nT_(z8?CFqtTJMf^0-w0LF5U7u;mFz+Kjnl!C&9E!~i`Nu+f~I&I<T!L^#x`T?!5TA*z*T9AWa$8C1=RYq(F#5xh@Oq0G8r0q#Z+#U`KV0-{>'
    '*(bM33pueXX47;m+Tmn-(IMb=z;|u-3z${<HWK-kAn7mg7KD*bmixX$-c%xR0eMXl)An`)gIa^E*{#~QxEGg`4Qz)^J6!>zD0Xn0f<)8DsT&vfP(^Ge'
    ';sA4-XM*F-gx`JgCW^f<45u^l6$m4^cUq7i@x?irf~krpP-)`K&q-gVN3TmKlc#HZ?Y4!9O4|ds>0>+Tb`e9{fdmKohl=;PGISK<&~Mr$a`fH&Zl6s+'
    'X^NfgBQ#b<Vd>}q%b;6%{~q=mY_#ocIY-@YsM^8fB1|aeC6hgfyor6#oru>Z@Ubi4B^JJbWuXggFZ)<VbOF$fa0&0)MHIcKAewAQ3EAZLJngOG?E(*h'
    'b$vJ+Pmrx)mm%DKzJdGg1`!u`ba_p?<sLZ9DsTn7iSR@pV3@k@!&Rk>&e1-<g?k@|Cv0lbPBth>pzKo?1FnPsHUa2S08ny*UN%DQ6SX}@4RY5Ab)Bdk'
    '?X&kF9<CDgYLqN7Ax9fygv^hTrLhus2z)yuJ26a^hDp16l<@eC4i0@J^mMULV80@72gw2i0@4dWCf*X8?~_&{lPXU?;h?mrufbKm>Xt|vgo!E0)3DtJ'
    '<aOZjpiCJ2?CBhY=mHOLx5vYFq4aWj$*Thgazh6RJ_G*K!IfWrd@rV2_t}E0;{?{8Kj#GKbhem*+(mmvoq-csNBMdTjs~}w^T1Sd4#~>blBzpCFv0_Q'
    'YB39{z`1YLJ`K*E8XO3=f8Ecv(XetFEvj@lbMExay`ODvW>HHy4OXSEwdcPUWVqcdoPJfE#%#PCsfzTA_h`bI%s_-h<5>_OK`2J`U_~3Te?23FN)YvF'
    'hzW3fyhH}{US&gwS|`i$(N+_r?=1=%cQA%sd%Z&LSagf`^BZ^qHu7TLHvy@jd(w0nn)@YQ*=_Lww2J!<M2F&T&TRMK`|26eHi_0JT6p`W3lzo0G^oJe'
    '|NQpj^Dke&Uk*U(Fn~%Kx|qpegEt_Wj=CkBmH7k>jdS>#Y<q&3*2A$@xJgYH5nSjCm{l?{@|8YdK;DXq1|*~z2fInMTnz{+BbE~oFdJq&q{Xhg%@l25'
    'Ag$q)k#-5>zPKRy3*_MbS66kEGHYKo70DobI4e68|BvzLl*4uT=hZ2=C{7?;{Ty8TVZR`>6Qiq@7-=+HZtpr+QBYZ71EIC&Er@xuM*x2SnDtfP(xMt~'
    '<Bg!bu3ji0_4j*ImqsuuC$`hc)#M5!n+0kABxCH?%mK}CwVGNCKHAsj1rTZgq2Y3i@`L6B9Pb&CK|tb#g9jh7|4_7W*^tnDGmA_h7#jMhRUL0HR3=Gv'
    '0-XTY3iq}sPSKc#`)FxYZ-P(_je8N!ZFcxrF+xV)dt66Rqz0ix3tQZ@+m1gyUQ2M<A`>)2@+y&=5%O?Lc9Gr=>FY#qN9e;ACt5$C6=en3j2fZEMTxY1'
    'dm|F-^x}Gim@kzU#q9_&{B{v>4I3Z(5c;{P(@sWEs}b@ITIvIBJ)}*mR*${<c0Z*Z!qm->x=GY-gc=!veD!d}Oq#g59wC<7Jpv5J9+gA@Xvz_Q$jk%1'
    '9<H28-`<SP8?PK3>JQATdGe1?(#C6!XV64TVEl*ywM&ygY<ft2cI!D=cviB9#HgmGHxGBc53hD;@;#z<&kWpK*|U`zp?>bXdY9<+2>rS9TBY+s&b}dS'
    'mwk0Lq%}fYmm{>i>pq;`byfQCdhEk_*Tg;7=aSY$r5cvZjBR@0>+R}odV4dZm*G2SM&AAmeHHh7X}o2zV1XAn8r$3TY{3?<i>GftuGHm_v`VD)Nm9KS'
    '(xsL5bdS8p!M4j82nOwn#mWLm%a0#XK<*K{rV-+ru7wj)EnG}Dvn4XB{yy_Y9Dt^AWNn*+=s2fKdD^}Msg-S+WC&d^n!qd_{Hx`gbvu1GfD?<5VVS+1'
    '%%-y&x{S%sJbT|}1_Rfpw&GecP0emO6heCjA-K{nRCPHZtYz67vR$&gG?}Ao2wp8#M{@lJu}Ssc9<F7SrY+kb2UEdDdYXvHIWKq)K&NbRPXVVW+BLw#'
    'Oqn6L6L6)@gZJ3#nOFG1D;gC&FN%)t$SaCE0n*E8Igg?H$D(2{GtRG7&6xBNk5#k^(@ih{sU`Zs@-@y<SKJM3mD=t%yZKFbP`X%_gDbJB8=VlEWb0xe'
    'dN-<-Jt7eFy%3Qgi=x6i6d9vCNxWi%!%=G|_qVRtdev=;zEYRj)ySWNy3<?WZ?**8>#hZ53btGywifZ+I-lfXi5}TbTTm@g=Q8ETFITWeFiKdSKY#yx'
    'RV;6m-t>2F-N+%fuC7lRGjC>#{lhGZ{AAVb!8BzhGsVs`RL`R~mPHj;7HHPKqA43vS3*rzj+|N~v~vKpN&t2M5N?e#<Uaa5OUt?o@?TI8Xmf-Bj6;9o'
    'nq*h74OthHweueB6@Qz-f=q#xPTCJ`v}nl{(wjbm%tydTSK0sTKbikeHv6}2W*|Vh1OC?N>QMB9Z9uDtH;^kmsSIgIyZL4~uJ^e8<G3gSU@$7>#jD{A'
    'g-Nv2>1^AP4h8%LwMh;d39iVY=gkD$Ae{n3f>1&rd;R+vyd8s9(Ca=!WeoP#VmSb5ssw_k5dD3~kjaTVX}7cDZa&$NG3a*FeSq&fs6VgvYxv0i`2meW'
    'Zd&JvQSUb0?F@CoHSKaZljzt*DTkTLMV$J}0Z`Kj5GxifAgt!dP0~EyfLga-w8vWWV4}2_I4S3wGCPL6{>yHKrX|AJ?H7yuBpxBzgrqX0G^*z@`x`-5'
    'Ng6hE(9f*EdIs}=pA4sNP8vcnwi2*c<B@eWjguI17u;8PXmQpTxm^z34t<bHWH!&Cb<u;dZ*S1CniG@V14g4%-))G(YC@q6d%M|_GC$pKHgotU*swv_'
    '`NywUecDFL=Cq2WzNg=R`Skxh73<%gzW@0A@rkZZ`0vH$)5jlAzyJE@Z%^L}7U=fBfBxm^mrvgZh_Xd}Nls?uS~Fa*Ijm5(x@oro6I7Vlbaw^k5dP)n'
    'hA{ErVm&~tgfBE~0O*$Mj_mxw#7^9?FE{oeY&thy(QzD1bToY?>r#Oh+#ZGzpNQ8Sjj+GNjRZF#?!k5gygDryumV@~Ie~4`a<hnj+b*^-nge*yUAG2a'
    'f|CNbBAmMo#v>S4Jc6=t(5^W2v^)c{PCUe(MR4T%!6>+xgFLgIFJ>rhY81cggDLx5z@->CUy<e%rsaP2(B8p!EQ))$-hsE?(oIHJ{aY003+t*rlo7MM'
    'ugOVy=8R3-<(%wIVcL%vuQQ?<_?pZI?BTz`*iSTbfpoauuXfQl;&4byyzw!Vq7P-|a`X*3#&;-UB#$G-Aq68|q5a1%#b!E0kj=x6!4n8H3Y1%p5b9d&'
    '3C@g}peD2VUD6DD(xXY{WuSIC%3^3a`mhIi30=*1n;8e>z*3(Oz7VHQIOZfyv>=~Ba^eH4Dxa)~bSJj^T^F>k-TXay@lk*p07U7-*g8ZJo&&g{AekYA'
    'k}*{FSqPM<24CUhsyJnpIl8n+k34x%`8xV{fto{eL*T3U0k^QgkHF9{5)tH4t@Snhi#|#U2&fgY@hI2NIHna>Xt96W|Mp4@u&|O%5bQS;P0eAK;x{}J'
    '%WwKXsumzmp-O8E{D8fLQV#jMC5jax*JRSq<o+8JZQ;Le2l2h8YU1P?SJI{(#^Ht7lDk!OwH^iGnOZB^vFwNpv(rvtdUkBh=VccLPOkmFtpIT1Emt<|'
    'ID+kt{3peX{H^P59*2HnAT&-w`q(`<B!ru5cVECIv?=}u|KH5-@WwsHNdq9uXwO)IpFzyAK$i#yDx7lX0hUW^){Gou2cv_{gWC@+gwyCLNG`{R!ouyn'
    'n%@kk+h@qtIpkpbeHC_EN^%nou#I3Z;mm$OD6#TRua910#$OWkY+x#7`J$JQG4aJOFfqUZ_d|26qTHg}Z_Bn8kWwb|6`7#oaopawlgT3;(xWY6l<*Ut'
    'AU5QnhST2Q)`rg@TsU&=a*Z)E$XXY(0gU+WLs%9?9L`tsa1+Pa&qLqn3xsijFi44vc<68g0h_k!Y7f#?7lYq!=Qs!w<I3V;_#tBKXe_b$mjUS!Af15p'
    'i-GJ2kWE1L#Xx=p$R{BGVxW2iC{4CMGuitch;$Ju;|V8bzD_e=l|P3%vJ7=vhD!e&>d2bZX-z8obEqQ=RHp^1{Li6|tX7>?t9t2b#T)6wQZ;F*>hdQb'
    'N7kuH>s0HXfE-z*CM{BJe*$u3g_^WNwf_mo<K=M@e%d%PYou6nNl#;`0}VgGrGc>r>o~|SAb}k5;z}g`rgpvPb-Ge-+vyyIvkq>Jqw85>v4EDeg{c4D'
    'cX(426MKD&`WPy7?d^~@y|=O>wC?&=ouW;QzotzZwG^UUV-Qp^N+x?d*R+lT=`bwz{rqMRSnXnWrOE>BvmAUFXJRmmwMoz<rS`0v(F1$Zu4#ZMOr^r8'
    '(|j4wOdM=>>;|V14D{nh#f_mb;wBUDVNWb(TZ|E%9odz(vj@1LzkU7w<MRM5Dm|oggrp4IPISAn)`VI=8moU9l8P|Q(Fi?K*o1eO$uz~Fj^LObL(wrb'
    '2IUW&4$vPl(x#ORH4HGrb-5g_MWbsOyfOACmRZ{q%iIXhRizE4#DR=_-=S=~+;66+O5Sxxa&}{*Z3Qsszy(Xjz>uNc!C+$61!jKy<JZezhXOebM5OHi'
    'GKy*l=}$tc^1vyZ`V5jvNY!a14af31B%P4z(@5?GNG2gQr;*U;^9)iW^Td|mz9EzJ=0fz#JZDN3uC$<W_3>QPwTv@m6hJv+g_!Sb*}a_?L0P_rQKEx}'
    'J*@{VX`kyE=G*pweT)NAr6+|Ow2#|4#_W=Hpn$KQ!}&1<?2+=<Z{G*hX*;lohXm_Dbi=M<c=4Ko2N|Fz`&-N>ipm9NMlm~^EILYdxM~aak6+>B{&_?f'
    'mr9kj9i!7kH>9$O%8gO+!XA;iMCM1y%qUr%$eK~I(J?kc)s$mYZiGrjw~TR?_R1McVM>SH4nd48Mx8lTGK0k}To50M=^k^%TXKXETJd%dd~*vL*x$O%'
    ';)cvi(TdT3hopSLK#3_GDed7+>&<LA-!C({0*%3QN~g2~%;-*5A@_HnK;F!v4w(gqomLxdNWx{nCWtB>hp)m}0%2u#H(yUNkaJI=Ud3kF2QW4s3}qZ2'
    'tIed{vHOk`PPrimgmR@QSdZvJx{>^wKEEF%;2^nA<T@kY?(vpc?%&c~b42G#iHmBon{3@7i(>QN^uC*Ryq$uBP&EhLAR27GvX56A`HY&_N)!^r`T4BZ'
    '3P8Mk4BIo&^!S1E133&v>~IuzfbNA+er!weXe-5eHeo?2<8{8CT}$d~G&Hs&%c{860m5hw-$OP`ehP76Zs4%4f(!x(v+4cghrt<8Dn^FXb*TqYNpPh^'
    '=k^0!1;uhl(wG?{SV1u55rPDR1ulNbkopl?2o(dyov*4RsOX)}2h)#o0tm4X!D##uL|B|eUk}OCbeQAm*tUx$OmPmNhu=`8>9EJsk)8U7>??+I8to7W'
    'qK`p>Y~34&?n6jWuF{&5;UzB`zG8=2JMeQNcYRsBohVGM?xQs6YsjTQ0UIIeYl)Zq6oZZyz|$V$g@rDTiwVdP9p=kk_5n>i-718CuTieP4}1j<kwmz|'
    '$BmyLCC;PD*~+DS+5PN2NeQ>0^pSHbAgwSlgXHZ?N}UCn5^n!X5`E^tOWqCPMZl@NeF3~HF6s=OWQ*Zv#{dxcY}(l;uZ4kWF)sC)D?&;Bu+r9K?(Zqj'
    '>}IxJAQQp64RE;{!=&pcQv-_dYt2TGaTHkhJJ^OGJZ`!^fQ!M%aPz65I+$WYVav|02Jd#p2q|6sIVs1%j+51itUe-pd+KdgSf?}1$dwvd@p%sHYT3$;'
    '<wZGRe6zB<mu@WfRgnVtpg5eR59Z~T(wsos3zZCCGhr9devR>xz8oM}oe-)_&L_bj3j6zImItKJ*^9xTeV`}@0}zc0Zj%Y;Mb92MUk2@;lX-W8-*H3X'
    'T+G)V>fxZgN?87dSi$*l1Xjr<tU290X&x*VU?kwW#q+X_dFGVr9+VDHKep@jbO2h(2rHY!4RJ%?bjT+Ueq1}Ys+Q|q*Cp?Q`D&K(HH56dvI04&7Fn?l'
    'eyHPA=eO+#xV|w-nIxs+qI%Oj{p<f5WI6#<v^OimuBbCIH$cy)1M(Pat3qNMIs%!N#9*Xn?yg`vO~uRCM=?^s#dEyfw0%n|Yr7HFvn|xMjqzOe&pv>S'
    '0=Y7V3?P|m$@10=h5VK-p5CfXZ~90sUJ+>j;_4nP39We<0lxZ(y{N0Q{PLw}3cS2HQgpb|+Ltm(>>nR7&**Z@JYz9}HPU+WmL3NHO0vNS%hD3OUDrm1'
    't-|?dXepyaqxOd6s~~1wD1T7?a#-;SkVL9CoNu_9y=7N&C!s3sdcL5OIBn<fRhENO6lguWD}{A^dWf}7h;%|n5KVcoWnt<9txtG;Vo(i%3v5*nV53~='
    '3|}Y5AtFoUj5AyWIP-a28|;bIKmV&HjbM2*#FeF}Nk)pb5Ga%?VQ{!^JB(eiM(hO2G*K5gWsY7^-ghuV4`Eb-sh@$-XJCxLa8|<^7<&fBB^dn-j6VZY'
    'B^dh*OnnkYiMbF|N0$O#LwLZXjq%vXouch(JG4X?B|4I2W6AFGx2NKlPeuFj(~qYgUrFpe6za>5;sULq{^OTEUo~39Rm(+0vg-yd@3VGO;HY5#(Z!F2'
    'Z&p;Bs!-^u?SqMd*>cM5i=4}Rpx?0F_>356q^J<wL)Dn20K^B_*CV2gF~a|?ZX{7d#ErrjOwrf@!vpR=zv+(B=B=?>G$*#q4qWrgcJUsKj?ph41=|<8'
    '=J5rw@KeZ%c;*EnalYaWHod_<|Ax#~`r>zZB|AG*P?&ArQ*8p=1l#R?7UGptP{9F<NBsWu$B%FQP)q>RvOBr1!bvbrsX7BByq4MSF;EhPB~iLv@2CzN'
    '29gc&q6+0584l0ku`qBJtx2C*kD?8((9<6((i%WDqfjVmEC}KRf)WHWbO!SjX~bM>j;btiAtF8CNWMucR;z)8vzony<4?w|z*E{(CviRPOIx5(ilu}7'
    'A4dkDEn9$2*?vr&<Tb9)ac-580EVi__S|~am8^b6(qs@uae%qewlFsw$rmzpHYb9y0Wk(XAYyeMF<8A$Lq#RRR$SrgG*q%qpGK<^nmUg*X?IHs7CZqZ'
    'ej3a1j%T1qgg=SYh(F6#jT=Lv3P`Znvk#ooFb*S<qOE;Bo^*w|a#YUg1J(3Ff$t8^L(0S=$JBb+zp8AhWD#oi=;Mt}r-$W5Vy$~BP<FrqK_w-scZ*?K'
    '9zw|qG@J>05~@1DUrL{WsuPqt2ZdqV<=Ig=DHX_BD&$%S!Izn50b}@YNRE$6`9sk+<XMr0KuM!35LCo1M>>E)1)T4WRQp5ae*N}6@OwT=7r%ac`sHi!'
    '^xNkzpML-G`QzuWmjh^(pvg3T<Syo7exb(7U=hX7gVgnT>UxA)3pF`PPE#KhjznsksP%bje@yL#x+c})RFSIozs)3xkGA1aaXD)xncbAUR0FF-mL}S>'
    'L-d)kwO)z<*2cV4IGlF-)ougd*<z)~MJ^u>;HwPYKNCa4teykz;fcTQ=eRSMq0q^g#BN=@SdHQdnQhlFrUkCZe9ILOB{X=^?%V0U;Ih3#ncI3RHfO%B'
    'W01_<9C<rd8A(W)G2%#Lw_a)&D{supgCdl~zUZ)zj***$?0*KiK7m|`Q&Ms0J$bB&Wxzb%h042Ku~=KI80CkzF8~+;pbPE@Ans^{&;-#(sAJ5!(Oh9O'
    '@U7eVWYckM6!M^1xJ*{<d$iIIF?_<PM=_9pafQ%P0G|N-GkPtaxB!C2{uuP7haol~^F@oX)m4*R2R*YsnQ_@9&p~Rz&cB!gQ^1Vvn|UA9#v-838CSo+'
    'k37AHmu@Kg%odZY=>RSY&lb@9h^>cVR=q?bVl^<ryDkMxTqt`$mG;O8<ZQIJ!z)N4%_Ia6FP8IEgyw{F;G__bwOsK1WhcNqceI?3C9@H@`3Ycu61Wmz'
    '<Hy&2((Op;z3I^Wb-eI69jkx-m-0&8_tC2~AG-FhA2#S4<HP0vMFQqTrI!KQY(7rdt^wL7=;rm%sAipnmg#V%`Uz-z8amqFoU59bg}MPpxW$+iT-5&L'
    'oO_K<U*P^XUsIoYOA~WD;Z7k#^=`l0=avOAbCC;pTeLUiS<NCG2Y^!6r*1e}FZ4p#V)__KwCGuVHy?(>xTcaNt0?0&-n)VAaArXpgQ+MX?gFEYh`|i5'
    'p;Vh0SUG9h%*6s$;80OXJ|?c^BN}&#*hXT|`ECn`(G-0YC3DL*>K@7#ID16$1qpWLklrSGe}+D|Mvll`BCpPo+c9z#E2&ycKs$c{W?@t(1T=uF65KQ9'
    'X3k8(WI>eN&OkN->7Kcwm-Fen*&}Y`9o;l1ASK$#I!8vnDcV)AQi_}dGRmbq-afr)WYTJ#CA><yHkpox$19-0m<lR?Q}&BAg0xpASUOa}+%w0o48GNa'
    '>DNfN!H7syjeEv85QBbBs^i|qjDF`i0D}1ut1dh-v$wNtp*gmZOte^5qmSZZw-gvtqfeu#YYJK%$B5cH?D5<hBQgW|3wesBo7ox#7y4RcA{~Dt99rBY'
    'aM=PA?ZYPl05ky`1ILI3wrh;54i>{Po0j)U@viIBMhme|gbJ5B7K(xY1{7~fR}Qv76{jEgbT0~;>MC^bshk?vr;_@p@Y9>;QWutM>S-cLkK>1=QqhR$'
    'J;y5t57u3~BX?_5jzVD&TimpbcJ1|c-w(B`BqhvO>@kQT12O-SoepV0lmpLkk5s6j9N@)#bzkYQa?3eaj*exscWBov*qt@DfR$-vOxi^sa4Nkk?O+~L'
    'GZ@BhPeO`(Ucnuw&s37y+|5F5P*{o?3cq%M6zk6U%8-$5fptni1h`vZz}^<5@y#5YW~5aeT=<oZW#0h8Ji|2G{SN0WY=hNZAE}aUP%Hn8Q53iNyHl*C'
    '5)UpQHKXaXm^y(|JevLr;zHz)a5nOx5F=n&n)6??_BY4_;6~hS=94|Qdl-CQBjxoP7Qg~%&?JvOlfQ9Fg4yv1qj{wcR<vip*ur`WEXMfx+U9A?lRb!k'
    'g{RU5D5`~1l=c<Y4;Z)vY7rBVLvV4QwzvoJEUF*s%7xI2o7QRDwAYix5TZ&DE{C|?Kg`DvZuflzDMsiUc1=o|r8!8mZP8tE(VAgkY%LBAtf@X)VIpL@'
    'i^|Ix+ymceT}5Le;J6o?zo;a_8J}of$GC~Atk?oHyJ^X**)OP*l>J~l82XZHQNsX<1m!aW#Fh+L{&vdMT*F|`jx(C)uokCq;yIfk2s4x)>ytH-zK$IW'
    'h?(?0E0zC(d#$)ADBY*%+uNElY77;|51!{E%h`52KgCCY3+RMISq9x^Gb{-Z%Zm2vBLZ<0-9Y505uH~*LZgo|NdHyYugQjZR*FA6^E_TXz#C{2jh0(P'
    'xghN>(DD-7%+bdMcWb{`f^AH68djdIqYJngSDCvwG^tvX698uC;*U?ieE#KY@#*^?pFVzm`fadeCbo&F6pOh>40}li=Lyzc8IemGDVkF0vUx_iz-3Qp'
    '<1$I(<O?-NUBuw1YcT{bJG{L6mf&C!M^B2ftb`R8-1O?Qm@NnIh#C>hVwJI7?n9FD=cBMNppKE@4R%mD%7f}hlH}zXI7-r2L%6utHMeW)74}GilBg`m'
    '^(qIa@Z-6n<p3OO9X8lR?|BK9Y8}3VJ-f?g*hM*rDQ|oX2(L8DX_3XJ#oZPs89k6Uk(>>b#}RB?=y)W2t?{)?7Du;+(3A~mUAJ%$A2~>MvpfvnDu-C}'
    '<v3P(FCMzFH&?Nad{c5}CIe)3<>1hwgNf`2ldJWBTo|mbO#SRR0rpFG$!b1(KRY3}HC2?l0rR;~E5zi+F(lf*oA1a%IC=y{9<02H_J%<?DJR_k;J^M8'
    't}T$lF{ISJsYXId8wqT!bnP+=zNBym>qTbKSBC9f|E1}QSRGmQn>rev@z!}Is5=en1(XA^Uk+NI1od(ZIV7sLal)399Vk2KXam+Uu&ya3;o%5}9Ovgi'
    'gk5_sBh2~Ha`#~A*`v(3HTEo{DvCkLeave7JcO{z!Fg+<zFHZUOT$zSoB~Oo2QdH#aH60g0Lo_fToZ%ro2%}9s4laZ-?hP}d7Wo%BjCl4r{8}2RD6M}'
    '@agxjmuO3%5c(PQp_upH<WlciP|b3m#wiJZve=_RYD~gMFq^<!;_hOFYS)W7+;7LhbS|Z6f<Wx60d2gbj$Rmzua)8J1joDB!_{5HMqfDC&FrL}3mCDC'
    'MT2_uAZHRavF&_8@yf88A@CqX9bj;Ix1&2RJ~Y5qTCm;mHV*Z|Pcl-V7x^F*-^P-2uq+c$m*3B@jn_5^m-7`?i&)NET9yL9c#ZmaQok0)Q#Jh?Wiy9e'
    'e{`9+QVN#?O>KtXKtV_+&!8)%z!6ecLu!dOJckfP20&=k5TPH!b{0nVCNnAADoDM*BWdO0<I^8ce*wAi@1MRs{pIgZAOHSzIY6ir0*L}~H=p1qBL!p4'
    '%|oF_^Rp1%3p{=2i{ZH~#gf7&Pk>uCKAIYPt&Up-?%U1g5i`AEBMpe7+rF|K>^WMr>~=r9>tdYSuxJ^8L_rcjh&ON(ve+4CkPa$goS|b`EUkaC`kpJr'
    'Mq6199_sCrm<<mbgsMi4oi-=Jnhj|;7`q#iViqA$KVAmKd#-0m@;OWRcfNj8s;ceR5eHmhi<Ud5NFTvaCKLWnG4qc+lNSM+{=;m0$)e?AiDvd<PyhhX'
    'W6VJ1Ib7u#rDXrF{|rrddh%9ZE^?qyN5#^z;M=a|%h&)?)`>&9ImfRhvs=DWeZ)puH&|c3z?igbACY~-h23pCxpK@AtR!1n@BOpiPjN<+;IswRp2J7q'
    '&=<Us4ojDDQigAG=}pzMZ;Ef>ci`m{s*{aOw<_aJ0oYAy(I7Fh+7`<tmyuiUHrxIjC>0$*wPh<?iowIMl`bq1_jBf}?L8<!8%f6e=YL^@AL;S9%a@aZ'
    'oPgPJ;SH|Liv^!ryj4e$e_L`NomxA%9^<$jH<C|j*b|UyNbTMB`b7=uNIPYujgC`XATG;xSnkI>R;t2cq}>C;q36o<)1h7Hnn#$f$`~Hn;<{-l@P|o('
    '+mwhiK!{PGwX56<rfbtp7BHYRQGK{N9-YQJ>G+IRnHhfF-F%waXHs+vm2c&yAmtFx3SOmNf#*))xrC=*ikJQ)Ao~-@v5`aLIov+?yoD-Gl&ZgJxZLsW'
    'klrSGH%`B?p+?CVxl82!1bKZu@n!V=je__Dl44P?^<>1jSLOA4ZOT#lCf>h5-;|>{UA7KIlQ{2E4FE(4^PEyQX|bomM(A(4URnm(39tu&cxmBstNk00'
    'm8U&}j>Cxj09`3=KYxCsqnx_hwo6i1^hwU3P0TlID#v#w=GNTK(8h&fl@8_!^LZ6qh_&g{O4+0))PX$}i5HBRFzt`jWICMYLJw|mEftCou{m$-9eg2v'
    '+K7(lK&*{lShJ-yV9u~$78dQ`iXXwLhTQ^lhsbO<Cqv=QK%D^@qZC+gpI?x7`%Q;R>Zyg0uF{D20)NsJyV%p#+|0-)cmh|-vp2P)+*{heEt(NA)}h^Y'
    'i_}o5z^0;u`4e%XsX6-dSXV_Rgnd>Y0dkG@oc`h*Dr~sdE!S{v_I>+4!22@6$Fym@iLk{jdMrRj4IxI7<3Vn4$_}^yRTpzBVL06|uNASaU>ozop@Uqp'
    'o8F%tv?-%3*whXC@mp?D4f6CBYgxnoX>UL;o~=T3>t=?$I5cZ$x$`~OLc%uF3L`)JZ;+d#=Egex?U0JUPvS)xbY^Njd1}`JtV^)f8CWa_u)2E<aFu|Y'
    'Gr;J5LUDWbp?d|oaP~%#hho0E<=M=dq_Za%B%66On^WEgTF|t41AO4Rm@ewR2x7F{dyOSl$k9dZVmaTC$SH>e2OAU>R&uVnIN7m{#Yn<BfO0D?tN?|Z'
    '0Db#3KHTB674?La9(siZt#R6ivN6RIcEQMk9aovk&1jhxi><U>UA^hYaUp)>;Y~H7EF@0>9tE7&bbUp`HXzsie()G*T(TYv|3udco`~D_1NJPz9*@Y='
    'VI#oxu$ofb8QORjXw)6R*aS1y8Hzm%1D^v*H)Ub`@LRdD7@@JBE0!TZlUFc1b1wz@Nh-vhg0tBUa5!<7a42kstcLH~zwS}N#_xi!?N-4%GbvQvudU!<'
    '(&O|fmdATlq+*5x6`cu59~{q$bs9|n&a^=%K`b`moO6(h8Q|1)X;l-aTp2I<4ag$z+x7afn9Mi3`<oVn<u>q*yEb+Lgi~fgE!Wr61&Anb`mET7+Smxo'
    'HWEH(OPsV7wN{PYcHaUIom1VOq&OHeK;u(OXqp8r5Hn^E_wxw{Ys6S~Vh^f`w+R*v)Uc^01Gq-uxPIVdznP#jqALoN5kPX=@1jNcuNV&k@7iJD>V6}~'
    'nQ2fe1R)VQ4(vkPe-Ae#r1$V)4tyc%J9$jwkg|@Hy<;lI(_=;;S$Kb5Wv1zMyB{Jq3E92~ITf0F9=%ox-H!-gqpx9M=@EI^Y|-63lDR$J#r><;hVUtR'
    'gK1v;hIq@LL3U-Z<za)VSPmCuuI+k;joLYEc|q~k!}apA(xA{c=PRCIlIb=iP;3zFl)nOCg&zXr(OG>e3WB515l}VO?XI}sniiKVUl^zR4BA#Iat&RP'
    'fVb_|)PjEK>j(N~Af{UurkFVT$w;DUv7Jg6GL6?1{nl>pDQmZw0#kboRZTeX=b-dJTD7%}Y$CMzJ;%&C%BDWW$U=lXL6LC6Qo_I-dQquE1zhHjZ(k3W'
    '*feLa^_xl!4sKJ&*)oRH-v&|wivoAEC0Q4Pm_Ynu|6#;}IS?2eS1%)cPE~vDgvRo9M>FcmG<1$okTg%eWP;t6<?L%oz`b5{?`JaP&>V2vKQH|kDK0=m'
    'FREhbAcZS(@fRwW4ARK-UGX>Ascb3gTVRU|taMWxH2(SJ$M?&4tqk5|mAt9R9R$yrPC?XIg(BLAdfbBVLfd2n=ZxLh^kcDWldv7)sh0cA-JXk84N+`D'
    '(F1GQ%Ta7X@pdqlEadz-lq#WAgRz7HSD}yMi{PBBbCWCvsTLUZD@?^XTr<)?6;R}2uq9>SNR8iY>H6L)WEgI^y^|yWIOaSK8(~L5Tmo_DK|;^!5sWHf'
    'ROc}y>Mp2nBXFS-A;EE~0CO+0rv#;VxH6$V2+^4GNR`fHd6wK=)S{1pd;;=EKzA?(?(G6!mq3cgZr}%vZMg_26ip^SD9+;yc|-daOOK=8mIVV2Fud#('
    'T~)^=&_XN3MSK0W#guatnYVd;B-F2M)4%^fZjdGCm}j2fc|m;dvD+6r_8CQ+1DeDi1{iX~SZ<>c;^wAQaD3z<#k=QmIWYM<cGD4zqfnN}&cp>qp^ZVG'
    'ZLeVoF;|C83vlgWfZjC;7MOo^KikgV)7@G;zU><k@iK^SVo5?G>KDW_$#oH;&G6zPJKeV*g&t_lE*S0^i0gkZC)g->P6aqa=ml|n3V#bX$H1rH*dRqP'
    'm10F|Gokx*B%Tx>uhZi;e_R{|GU7%LXnlHBm36v`!LakGbU7@jB0H-eLy$CpUPi(wiam+aB$T|cUN4MlcU7JWrEpH9OF*bjBj}R|q0c2CxRN*xkmRgl'
    'g+_zxC?ljS`;#ogS4IGp8ZE=WU%X@PIU1&q9-sulwZa%8R3;IsV!oJaJ-AOo{|O}K#yxRLQkmh^Jly?R6d}mTZ|re=8K!s1V7OO|bZ6tyEce^#et{0t'
    '7omw*>rA7Ld048#aLr+1+d|ocf*mx#agpOVPWqA)Kp~+Vrx>(mGZ123El0p{vrTf*9@~*}{mz1VcQ*9K)v^~sk1o%khUmbTAew}z&maa3_a%rnA(}IY'
    'unupjQrb&E#fF3Sh|{2~GlMS1Kd?gov&d-zjCu~)zXG|@k@Yw(OcG`E?O*pTHgOy?P+o9f-j4e7w*LnS<>~=t^e@HazG<!p@2F*udiFJQvxRNYy5L*z'
    'PQF+D7GU_WOlVCP73<qev>s39i&@6Zii>}L`u6$j_aa0q_m7*`G6e7P)W*!0vc!*_&L_zuGZewR5f<P%=>vIWV33n`@;E+Q)d>y>a4x~QKHLnAD%=2>'
    '?P=H8U1GM|%`so+1fIx6zN(+g-(1X>Z}5+k>aq_7@e!_mLC^)wVP;Jn3%2l1n6hUlMvAu2qS|2r;O|tM^ynj96M0r;)bkO}!I_ZCNKyVH#qW;}(2LGm'
    '8U1{s0cNhu-on4c7|%I_pF*}#RHz)Ot5_rQWcQa-ePj@)NjUlxPVCrn2FIy{V^83qU)+||;pTjQ8^p5-uR1lZ>R_Jxz9!8R?ZE6*h&sPbqVT{7eM%nj'
    'TqA&RYx|5<e9=z!ll|$V(Wp4aTDki9T%C|GkC{3Hpek;+TAM81C0sj_je(X_36BQZVi=4AKurRw6Hu^fz7(el*?ojlnP>Muj`Mh~t5T$r+Hx)2c6CSf'
    'Z}0Zg`#IV+w>VX#=Q-H(HA%+gmeT$5KCaTD4b?1z4T^oDZwCzu7+RQ0e)chp=s7i4`q{KgsvJ1uQugeKoMFSPP_LPE(P(CPD8_>BT}Y}8u+2+hDGTs4'
    'q!CCup)Z0=ttL*PHVL(QDQZ8&_Y`&r{{-xia~8H<5*mjY-f$8y#k)RA@dhy2e~s49v-fyVM_F!_ZkvWrR}|S(1n;_i#N7leCOU{%wmDI+1Tak+AlIb;'
    'o`1c=kHrms*Y(1A$5I7OsbG~8j@;PXEXVb)5;GVyy)woRadg5-om;om24^!18S*KN&JCPtRVBvfrXfez8W-VxulG(0tp@J?IGgS`kbzzTu;%87DO8a~'
    'sgQC2_x^m%QHC698TMguv5dtSF<KZN8&o%Np~n(oDY|TSu%AEB_L$B#D+;k<JC7OYg;f(RP%1>GQ@MsvriK%A$(@W1n2JJO+49SeAwrc9T(6Dau+|N7'
    'TPh!p%_8}X_OL)-kCUR<y@6Ji#SBxkSPl<yq8D{>yI-+W?Xu7Zbwzzk26K#?GQw%@n-#qFC+<27`az>EdYRCpzdjWRSu6Qy#ijp-)tcYW*O<w>e;7+L'
    'sZBIoE5)7W0+XP8?Jkr4*i$kfj{6bQbj#69vG@_iyEph1lHRZ*i^sb)>^WUvJA!&d8;2_)zQC|~bFiY=L<9?-+@CzWwT;$ZW!V2dFoILsV{pnmXA7@g'
    'yz-QBjX43v$$-ydal|)adDquUiP_`vLyyjE)h=mxy%4Kzi_Q~D>l+Uzt3@6q(JD7R^_OIW<e`6RvXJNjn4D|V7=Il9yQ97tiF;lA`&Cr<j?O~pVE|aR'
    'EbZxmZDCetJiHUxq-fGm3Cq6{%M7rD0b%p7OczasTetV^JGeYTgkW@fpY#IPRWxTPt96-~A){X*1r<HU*uD#iYCN90I+`_A@YjcUs#)x=^B(p`KQ6NT'
    '6Lf^XxDuUo1F)uyV3tIMyf{7zxd|qfNkYdfEbf=DzkK@b>oKCNtRoLcxZrY|D?9l&lV~4&^QcRQ{YXtyl>@w{tP@^M;_8-CPSVI?A%PxYQyL-U8!1n!'
    '=A!F)F`5V#qg`#k8JBBR4C+vh<-kdZ>BP#D$j33#;4Q+EM$y0b8!fBmIO3Yh*-+8?dz4PHdNHA(S8yW3<XtS=>3w&ZTFUjI6Z^(13-|0VU%&nSwfN)d'
    '+f%Xo`t8S4vHtw~$G?9nZa@9@`9F&3pI`p??Gv%uV*BIskB@&Z?!G<!{>!H!tWU5$ix<xCC^}+r*7<}5hgr|<JP5{Yj7y^>i8R25z@^@nBM15Pz7N19'
    'R}jD}GU^fk6~WVz!m*?NfL&meu~Co&?+vX-jgxZ^pypMep*PeKqA3$%bq*1uLfRf=jPN}I_0a{Re5GnBJ9#_5+lPn)G=~N^J&0XwYQ#?1s!`mqBaC3s'
    'smy`08^>Fz-c=Dy@@(yj&4XwojyDhjsOT~QDBbHt6uS{_*ddu$N!Y2_mdLzuW2m)=sFm%ZW_5sq?ehe8Z$gh~gzz-RaD&+>-6k7ZR4<1Bl>kuG#Ctqg'
    'PlYLNxemk4=cg}EzyJ9B`%^Iur5U!NG{XQX+5-{ECZy@t-y_!lET*4-|MB#_SbzKc<>}u)eS69we5vIm)!aiUP6*Y<SkBup9UN`MOfybRe`Y`o5gt-='
    ')wAd72rSr4G1~6am+5+Ms>}<#=rBkbBSAG)JnuV9yi(DoOl`AAQ#!H=*pEf@zHrF)0u%Q#K!lS<zK+Yp4<b^xv6(Q%pz^WH5Nu<ll1!y4vhK%8?rz~K'
    'qmo1$swv3(cR9oM;{_H%@JE{%WXIAaNwje<k(NJnaOv`~D=wNhsvN??r<G-_uK@eX=Uw}{1&3*}AucJ1(h3mx1xE(nw+NG9wSil%Ow(s2@Rj<MuIG3&'
    '?Jxr`i!5?w6WWclyOSM**n|R$(_ne~n-)}W5XCSj8xC6+yE`fk!g2ZQx61-R-Y`lJ*2dmM)^qug5!fg_F{D4BRg~e7I>IU3@7d1aerKFn*pr%2iR<+!'
    '1<Kx@;;@+0lv)0pQnBp>Os$UoJY+JYe@0IUrd<DDUx9tZoa?mEpe)C#fi*X3n!tgic+?3N!V9K@*>b%HPK~jy(Zrc8@}$heTwN%#WGZUXK$E)Xrgym='
    'BicAn)9N-iTt+Tg0+o$H$f#arU3Nj_0~KvLE9T2Ja97|Q*Y@47*d$5RAj1Q2bpDN`XksnQHcQCFdinm4lU0?J-LJ^)$k!{jPJ%)1H%nBcI6(Vmc86gr'
    '<oTIyFnW1`c1gH?(QI~{cy3%Qzii5ybapC#wnI}J8$!25iKYiOmBU&KW#g<=<ZH63RS)}zp5PgB5z^7Ypk}cxaxQwWu;4xB1a7d-L7UYR5^7&6DW9s%'
    'XnEOC9ttbOW_&lCYf=xn{^Pn^VNPjCEe-n;woWQzTBWKKWmyxSrp_z3-DZj2A|)Hn7x-*OCmnG;x98*PO!HNwh}Gmt@I&G%5v%x#Wyr#YJ7I#k>}0o$'
    'uLORCk0vsc9A#Xn8(S(;<#P&kNySrkXnWrjH`u5EZLD|uHlf&9562j;CyWPWMxOlDTy|BIFT-FSie&w(Kx^<aeolT}&13jYNc)H%VPN90WN)N3l7kPl'
    '!Ky_TteZ_c#VHSrPY$}q#Dg-?`<OR^^c1MeNfTWBG>YvyBZwU}kbtUrrv}~pc=oJKA69d4VC<qN$OS$YNU+^Bi`&??mm}KOArO@NaC7h+h;`_QpJIYH'
    'VDF2INE?|VrepAzAyPxmDPV>xs&^TiW=bHP5Ggh}u#1SxZv?S`TYrXZ6{=8!Y~11p9AFHfMsw(Kf(nglsbL2fmJIKo0H-9~9;35xvEJ@eaRvpK2M34q'
    'z@4b^5{GDSibCsRy@wkdBjbo=<m5h)SI?8Pbt|W@61{((9^*tKbt?G}LKWSHRJBmbaV+Pc&Fy#d?J}7*FdbZ%*S@?eC7Bt3te78y*|`Tay$%Nh5EaHv'
    ';2w#HNe@OGop9_~oTHItIF(XWth3#4h$x<K>W31estc^j-f>;Tb|VR$)aWZQp@f(kx{!=M;fT=C&o|He4rj%>gL6V|C99zc9bvOd$P~NFNm(uZE6f4p'
    'HiY{Op-u>TfDq>(5odHwqO$|KP<L^ROk89QhlS*1+?#8Rs1hPI9An?nHAdYeYBp44)Y18NOjr9xYAZF)hC*0_WPFXynWkSdnrOzUL&e==8siSuC7Py;'
    'y<S1lqxX;tk2*tScTI361S*Y{A(K$0&InafT|1ImD<yb=SrSjWo9|*a>$JU%g7S5ehW<hU3D^Su*Kbc>L|N62nFF;>VS~_#ddE}HTHZ(gqP_Lf(}vr4'
    'o*=uw;+seaV=1N@7)z)g+@ggQZu!R*GU(e|G!>E0Gkvo6$;isZ>Ene$44v+3s!5LGEPZ68Y$V7ESx0rmI{Nh}*c6gB5z3(^$)POLtf_GJY1y+v(Q{1l'
    'G#E@K(wHw#)@zKch|DM#aKm7`n#&BONitpYEJ|boM?h66t@2=UY`<gaw%cQ^txfTvUAI(>8{4n~Fj{E?K<(pBbfGskDm`!8={qb+$~C1vK7IT21ON8N'
    '(`BXxWe_1%xokM<n5!}ksytCVM!Ut*09>*P{`m9b-y^(yQeBI)y;0o86+<0j!HKoTIuog0uW3vb8F0ndH1&(7YI(QwEw>ZM4?pqb3rdMOuz&I!epXlf'
    'H`?@q>!l-zt7we^u7HvoNnB+|2$+W<aU3;yKxEHK-hv30c@R3a@4WDsaCVu%QiBqm^^~40C{$*B)fT!Et-*r)AdSjdD^2a{x0vn%S}ptxRgm#UdOK0N'
    'u|jWBV=Aukz2;tFtJ^t;-=i+k7Ij8pr4yBD><Gn%+Xi}-k@)ly`linxzoLQd|Ni{r)3?u0#pd&WU{5<q_XQHC(uUWk-_oSFSeG~*U3mTIMGUcmJqmj4'
    'YA)3j13L-vZ%6u^jZ!5phw^Svs#G>brbuHVuv8bdTw@Xqwjxb)*^Ev57)o8syWsYd&JyB2d!`+@1`Xpf0Z3?L!+k{iSN^t@3YWt>7`5IMq(i6sq}Cg%'
    '3r0$(W2I>jM){^d=1#^g3?1nf@6i>z#(cc}WD#VKcli1n;K%~g#Ph9`w(cd3YD}RWC?yF&3g9WOE^&mAxI;owFpcIx)v040HqlVA^g72tzX-_!yD<+s'
    'NAe<aVCm`lM9&+1d7_Bz8_|H^&KnRi!z~cT@MopCoM^8Fx>(i|$DXW(HZ>5b)KM9utbqtT2TX07-T2mfBd+tN&KRhesq@*&eDW9SFKB%JOP#o1t)pu0'
    '8>M1Z6N#ZYx|Fauxsoz)%vwf@lMh@7SVcU!GRrf2O)7coY=s&nrC}<*2+!;eL4f^H6$a$P3*#Pa)*sO})<nTaH!5$92m(p_xL;A9olOUV+J?%3gz61f'
    '7u>Ma=U6V)fMdmSSPPL24Ps$%Whu!-6?f;m*>vk|L7lI+(c;KzXR$@U(Yr({nI{$ZK^Z5_>PQa&tT;Q22&Xdrp@rZKJtxtg?Ec~o=E_lbDxJaj1k;>>'
    ';gOs`sY>}8&WHylxd(X+8(z?!n%!XsRtxgtmoHI<jMS=33bZSqVHq*It<TIZRDHnu_m_exg*QOCuUH!k%~8U9^kZy5-Yc^Oq1&97V%ja?L*~GJx3VL*'
    'p<!o6>=JQ3PR!#SlIk)!z76--<tq9y5h^6c$@5B_FIKB8xq5I%lfj5v1O*p8rG56V?xWf_&iLfpVY|Al1zH<1^dy~cJ{8?2^b3m^R;SoGZaEb6A6zwI'
    'ND-$WwC1c8VkL=YeI8Dipf2+_rr0ROy(q?sFfu&R#Bgn-X~t-H%Mad^;@{XYbc)LNsZ63$r>S}a6g>*Al`3oG9Cc^VnTvOKnS+6qX<p<h>^P2v!bG6b'
    'Ww~#dB;k{Wrk<~;Kg`J&phi{c<ocIf1}VP`gSeZt*CYsHg?)xZJWSDN##8iJ5*Go$y98i+0KvF(M4nOGDg_+UNQfHFGCUVui}~|T$z8+I-pDFZ#9Vf0'
    'oJqpLfVG4yiV(z84&}qxyi`07iLQ|_+bSy#C~oAMxGO0v(H61A${T3xju?&2&Y}<DpJW}t7|;;=c}vN{x=3CiS8ExxYJ65Nh{WQ_{KK@H-V;9!?Hn(@'
    '|M>dxAHslZBPI#ZKa=S#RJpE7O3?WD<*Mhe2BqJMc)<d~;G5&c1o9Zoy&u{id}(7yfyTJ>f!;xiV_H>TdNFJ@tOX8leGVXv&vS$$O^YH2$nA`r7i8VT'
    '44~QHuyJ0Y;!W+V%sI$eIS;X2x-K{qX(mSTO8yrcn$@&5`NU)N0oFb*-nUqA9t{Vecg^9p5C=4N_ywgHJKn{$2$`TXg3^q1ac7w*bc1&=rbt*vSsrzq'
    'QcNdY+2Zzp-iFu(SOkNzEr#p`(#cXu3Jc7W8c9H<{A~>C0N#S`tc+5&4p7=-%J(haaf>#+%vgy*1rmKCW1rs$jL~>@d9C)7Spu=9Nh2<sV-TX@DTAep'
    '+S9-jRF^OC1gQ~6kc!>ssKyOxAsY;irLbvi42b4JI;p6RG7W||NbL{~bcjjNar?sLqIn616G*s<4-J;L1L_GbCP1YkcAPzZp<vHfE{qXoh^id=ple{+'
    '`49z;eP-Yi11|qqv~3cTtx{eJlyj~Q#~N&q+X4@u6N#xG%UxItxXrDUBz8&<L&QSW<=C?1pZ^6qtIJ}-lIy^gWsC?z$xNM;v+2$tH9r@#;qyQlsGF*6'
    'MA8qupJSXip&H+XF8=uZ=|7&nUl#Yv3Y|GQKm!70C50Qb;RZ$f3DzRJ3$;n_mzMYglPTo7M#5G|LCB54yroFI3ohK?wTb%eAvRTmORKDsA=UNmpL`;&'
    '#SEIlQEs}t5(U=UEv6xk&uJa)=}qv=4E%1aIil8caK&^1;?J8<5ivTo5ko4uS`M)t!HBJ~B8oXiX%dP$fx@ztfY<>B&UGnu*4gazlDUqi@T^@gN}}lu'
    'o93h>x|zXs{~`9e6=Erb7iztp5EDnK8zTa3%?;vNpC3eQpcuv!URS_@lo+76?G@||*81-#@~T_S$=a{rP=N*xL~r09Vjz-tz#B?T3XxzFPR$(LitXY<'
    'M_xXvl!-%$Hx^vQ|M)d<mr5IFL8?^JX-D}|i(E&Eo+L!@EChp`wp=c_E-02AC02$5eq*AK9Jv?=*ic|3rBPi|6USU|6B)V2%C$H-(Qc31tC6$QueHhr'
    '|Fj6tn8498ErMZGUff=a(`UrU9b$y?_Pf~}p@;<cctQcj;>+4wan^IwjM-ujgQEJRXe31EYYErPW<SBkX$lUBtuHeaY?NtS=r4#YobKD*euDx1D;O7*'
    'p`pguA^t+>KrRfKEL9o4t6ki-n4tWgB2Q*pWFQ-0=sxA&M#Vv}RO%;V6t_nk8vfbP#U06WLw_LGR@8exMh;`4uc5~2LwaE0HZoU5V#=0gZo1#R=O}+5'
    '2Ym?Zqn6e6{TAPus8#Ib4Xo8wA;%n#7vBeA1X;iu((8;~*+gF-(Q_>QfZRlKWmN1P27{*<qWkVc-f5L7)AKrd9Z=(<Gma^;$Q`c>49X{#7RkH(#Y_C|'
    'H`@p${th^DDP|CDqrf6|t+=_nEleb>omL=LMy*0oi|hyj{6CcO0p2!^8P}U`3V&N;wKNQG3ms62(mF{H!!oG-;7KHb8+EoR2rn)_CfdP0G)3c+PJ0=9'
    'M@OuXEi?K0<LmDqKmYhzY(D++_4i*s{e}q#-+w%P{QK8p`Si#4Cr&B9OP*y#N#F4-@-7aB;=Hd*N`wvC7KKJ#F!g0Zss)KG1ZR=Bo}9qc8BCL4%I9I;'
    '_J-3WoT@rFiLKG;e$M2>XrTh-4<l@f3+txKOLV{|-;AQ1!3;`OuY^h=LxiPrtooH$L90(_Iz#iX9a$EQL|`TdYhDR<2dV}x>n!3`N`e76+6{$U%-IV4'
    'Bv9o1?Fqx}zvEfa#rGedKmOy>x61_JVz41HiY5i?^<oV<Oi#F5rL|Q3*O?|YZC5L}f8nF4WHGRKna0YLv5SGPz0_-8&Cx1y@#*^?Prtwz-xQy||M>Rl'
    'x8FYhpnv-1>DMQ&wz#`<PS{)Jd(&uNyTZtT5|+((%WIcJqQ%fB>e9LLquWkIPSO%xKqEcz67SIzp}=sLqNU(J=X|7mH9Ut<FwSMdOdQ4gS7UN?FM}rj'
    'Qdf00rKUOyyPcDn8B6CAvgWBBoawr1a!B`jNI+yoK_EHggO}a^#29Hp@knNfIX|>N_uV6~IX^T0=U_=xN{~5I)9U9S`@F60@ivk<egd<1p6dK)nx90>'
    'Y=?yD@=>~<U!#O5b{%f({VZOz8wb1#2I-&m+ULX6KR3+5RaV#eeQne^R`<eVS?4F++Mhx^y467D%g3pI3N@E5>Lx#!_UAslw=<jkNNWF6fCF~a^p0fn'
    'Q;>rb+2kkE{@f^sd#K6Jq^^Dnv&V}#rXd7EMpGlIN1zm96hms{N&44iaktD)isEdgD8hK+J0X}$yrX2g2Av;=GF_60KACh=Rc0Xgrx?^5zF^T1FuOy)'
    'z-qsnp(cUdSWvLRrq8G>B+Xlae3Q?%`^jA^j71rr^8jmp3NQ#10aJ&j3@kUDmmvniPE6!+2`c3y*yc4*u_$=LGO^ttB`VnJL>!NJKFQnQqH!?({IJJ_'
    'p6%fuzhdO+mu#14YjZ}W>z|s--c6!IS4`&n<qH4<P*40slzfrPpF5*HM3<j>Q~&G@C+t%@mEW1>C*_*KPOQoV>0Y%H*=C$J4>xcKNY!hQg3ULFO2@<L'
    '*FZ_aeGZlmh;y$2<3#%$sYzR;ejU=bcW5h>gkEP}gA}doIhIOTuNsv&&T}xGz|?ENVhMp9%OouS>gghRB!jDyW3=J=>3ejqDzqXI2)jmM-$$<~_kFZh'
    'T0mNx%*yCe&ztRG>~d_;!q#dmK9%t<xMg!~_qFzno|vIwa`Us?{G?!+Gn`N=00G&XL2ZZbW*tCo=F@&mc3k>5#iO<g`{h~`fDt6>kD!a%&{s=-#;pF}'
    'xZGlIm!&+fCBcjXXDIdPW)|GAcW5!iwyuS;(<QulvQ)alDl9~rUaOms3n)MD7IUhfHd__D{bbIixjG>*a+o;uV!gqV83Ttfl^4WyZztHU2rzgAO;zv^'
    'OqqsY>hvY`v2*#Mumx9>zvkeE38s{Qr~?Z_CG!CL?yfR#O5<DQ;tT5Z0-y0D`<kP{ruK9!Fx@4MwR-;R-?s6?YaSN>$hefwWgV|xxyO4qGl2CDjDICq'
    'E}a^kZUjya<9ecKz~0<Zp0p|-lt+(!u-raiP!v<j=b7YAuh@ge-tGAW*AfVpxd%@4L#E77WS?{S$YYcq4Bc*U(3CU!%@(Ar>qli9`c(w*wEIm`W9xBo'
    'F0z5+QqQty;B6O-m20!cG)>wnBY_tWbqe3an$!<UJ)nW9(A9%hg&bPXhZf*u*gjoCRa1}7N$iwI$y?!?esjUJy<z%v2L<23@%+*a!CwiF56IJjX#BH7'
    'k?1vCAzT*Nw%1{;;AeP6IwDn_?L_@7JPmX*As8U-ja14|+a`1jKT0S7dHcRik?;o_63LLdZ<hHvku}|b0oH1r&IpH#oZ}}CrkKt*RO%ik5PcDE;3fFA'
    '2y~f1v#TMCyXUYXL5;(`ZMPKa8gOjF$yhz50h}G&Z97GnmPS_TH_Q<bL;$^w)x&@(m+o{%04teSREOgCKppxJ@Lr%i-VhI*s;F|)Not@JBOIhz(&O6_'
    '^H2kp3XR@PY3q#fhC=Nk<b*<t7fcbq!|6sRYJHOW9)q{u!S1?X8lgA3Oh)Zdi913`uGbp9pY3wp-mS<`ju?U$^6;<<0tZ1;g2Ns*a`8wt%cr6C2tOO?'
    'Jhq3(>ZCMhtDVF<F!n^RG&hu)wp$LU27b7wTBP_)cJjPx(k3_K>?)Ix0$ZTbBB_j4_iMB0v!m6zxTwl6UxeBv`M5T9N+`)DJ#@#yeh;R4P_;oO9AtpU'
    'sE!jGL^MGTAz2hQ=3|Ql_^J>$hTP<V!>VW>7E(>VXSf2Gq|vEkm=QknnC>9K92(LAHfeX{ID6sfwmqXK#x%t0O&D~|Tm`cqQt^Wrkbh8#NJ~|{keV4='
    'j{=KCIrT=e4ye~B_u5J1v}B3`Dyp$UW7w2)PU*v&?gPrJCZm=5-VH~ngz(i3bz^;ut+j?~Nz5o93@h)~FMGeuXm!=DqFbI9SGsp{&DCR>g0%dw;of5V'
    'h73+ppc_T8yD<?<+}OnQ7V~Z3d>7KlE%Lz246T+NnJS;00ou`pRZ+=H2@a>9p``?{nTJ_wR|FbgMk6`=mMbL&#=4n3W?iQ^59E5V@4Q6P8%ti#B^o7H'
    'F~~{soMFOH<xF%wOa10*{V+MG*o2ONSk)%9Xo{KS_{V5SKOlyPCp<`@$bMYS&^)pmS<t{VD4;ckJyj-oHA@?W5w7Hgm0nla$wrPghwM-p)?G`NJ(9~>'
    'nsh62@3s(B+iOq0-%uYauFH4-SnNJ%4sAg43@Vu;vTJD+DmyigiwE}xbioj}3n&?4l9}pf2luuMLvk4aPn5(|odFQPEK((eM&`%5XNC|?-r#$hY=)Y|'
    '>Yp1<=G8_z!Kfu9TCrhT_7T@$hD#pPSdaHoOjlzJ!Q;k={DD<oR3N)Cs6=E*g?3_ebtzv}`^>1VV%H#LUwQNNZnHYHWoAtEvzR$6z~UUj11nBSqmISg'
    '&5!qD)Gs1UNW&q~qm=p<c?MYsXeBGd1yjU{fonKfCMZ5TJj|<ZF_aKk$!ZJCgUicbFj1~|G@)-K+cC4`Nj&bMFevpB4MpWKV1<?l31uXytg5L@-ohJ`'
    '_qy2bIz3&>)8yywKMnHseV3kX(zD0Nn8K5u^<{dtIeK=DErQasDvfw#4BOdeJ(n8)L|_tlbwmI>F+G`0hB<n2H^19ts6L?{Kf9Q9tMq7E?TRJacW6#q'
    'V`@FtyZzVyyP_c4RM7cvsws3`)4BtBZ{1Cu;9`<L*Imz*r}?^zU4k{GK*Yo_RQ?T2B4gKe+6s~}*KltCKpzt$j}gC76RomX8cK%_fi30oP#iZGg9~L}'
    'RUJ!%E3_o1E1~p1E0pHP7$DPmD4U;!>~Wh$7TI%_CbQIBGvFFSp}l;p!E7BCaVrI-N_vXTgd&!y5QSR$b+OxPU%H~2FMC+G_bh%fpe$SM(h{(7Hx^!P'
    'H|rTStpafrOHsib+(eW5f&P+CTvn2lmiyaVyfUsi!7-TOA^{c@S#`_#w9Rl_f~kACPSdV?H^;!{Pzgd>4aQ0Wft;WQzGrUv8hp=O6s={Ca}iw56SRZ0'
    ')12BQYCA%0demk}?Gm*cp>{oLH>9o;wI89bd(?H0+Gv&Hi%Fsk;mI7(_YO8W>7X2!TxE1Rg&VR<yWQN*DN8S9%G}?jU}g-_bICF}Zi2=75`7aa?!NZD'
    'S89sF9f4#liGoa$_m-R%aBl%rdyA=RWx~4Phc6RhvSpIEaYP6k7c{pIk+cy~l{re-t1-0hHymZxGws`C+Mg_ysvE*p;xED9xewXvSK7OKj#-7-1Ib13'
    'D!)D@X(bl{#qEopKu~?HkI{~$yI6CL>ItHPlIac%D)?f^aGW;zB7Bx@Pr}Q_%WlD@en!Lvse>2X__SB*Dp`cBNxGio-AT$bRj#H2o03jNXn>H=>Le|X'
    'hgya8pcLO=9Ze?O3x>$xQaGvAFNKrJI0PVIO_d3kWLn`iKO_mYODYb5dCYw6Q))36%D7<S2y_(A<Br35NOkM+)#-H8<E*MIQ(hdG^BUx?OT|mWIV@Zk'
    'JCUc{3!`38HFFfBj)iHsSkB2I)F|n?y%6S=ywy2EoCADu|Gg$w%Tf&|uLsWjfx1%+ZbPL)3O4EAzTXNvSU@P%WSde0vYhY>29VV=VkTE*ZrnHzmMPE|'
    '3?hIw9YHQb|Gag8RR#~AFDR+at>=8~e&woZXm1fJTy|s>$wVzO7=s|J-#8jzZPNg2pDrh3M~d?V`zdAg7OdI7DLyP_Z>GyPrD8nqGQ8x5GyW`|IQWKW'
    'x${lw5~~Q4@lJk*?@ql&GnXekl^ij?@<%3pDI<($l5e~Vrs}^o8C!kjPN8sPpA5xTeUvI&e|rjr)SMips{1HSax#p^&tN5#yal}$C^329?`&e~lM{pG'
    ';yD$RoiYs5pB<HI;`&tWVb$TXP~WXs0p+6Prw7z2<d|$pLuwewES7<a5dB@D%akBkAAbusc#|Fxw&ZNV3#ODA32rgVCDyTU1a(WmLm3JqW6R?Bbk1rF'
    '_vWN?7rH3L)_3ZvE|<qTMG+7JSV+!dozlKQr!=Q@O2U#)@;*67XuWrHWcr&K_e@+)f@^^9BDug*bHT^Ar571O*G{Uaa6y&On}-ymX!(!6`U@uLEmZBQ'
    '1W=y@5N|SLR0E9EtW~>{7{KL%=*1wyKtVq!6F_wmAeNxLV3I^GDa(di40ETXY4^8D_gHN2L1eqVWvAy1)+)xDD6}e6uvZwck$q>%#3k9Pv!ZGfBK}bo'
    'w4R*CkU@H(D}qq+X^JVEEX3$B(VoltQUWYjSR<m#1d!<lW*i{2gUJD!1dzh>DzadYYb8fxP%_sh^-9r}HXMj>TPlSWvimZx1nt#|6eOe*43k0)*iyt9'
    '&7_xa;O_P114{gGh!&v3YZ6sWvNS7@fs^@w)Xc&9aM8;Vahi*Y(&_+4Z*_ca5@z!XOeudIam@fX+e_suaN}7CsQtZEC8KYpUy{Kf;lUv*^$l)*t-KB|'
    '0Knl3^1$JrOOoJsMaZU3HcfKH1osk)W_|BcWN^0U8SHg8#mXSe*P<^_Svbt?2jQ<?=Q`UNdw+gY;|{?gcqO1(-E9)6e}yAKjA(aw`Cl+t+{V_LOkpnv'
    'X3&vQ#%9oYz+(|`0%V;<7?&@~n9bTCigK9Hl7?EoNbjE)i|+lbc<lDvO9(x{f|l9%N%rKEaNL>*<>OVf^as0Hj6dP#*fX9?K$OX>ld|5F&!uc~7oYd-'
    'f^3%PhHW1MkYnFo*l{}Baq|r&b+2OIMC@4*Ok@VlUG<T3H`6@YX)NbVtqziI%z0A#IhIX^I}XH%`wnAu^_3}?yKQii7Ni;=;bj6zxs{v@_$-JIAaw?k'
    'a!UF7e-@+)AaboZQ)LO?F7s1M>xdM=M^_y49HbKr&gSRM;nJ1MEwy*;ZjYX)+jg-fM@iprjD6eT-+>0SZvT2oaU%qnz^laDecE|(m{iO+*qR>-zdsfZ'
    '-(Yh^)fh&uj3C!pOrQ3nu`bhVi6pxTtw%E@i*H%r#-+7%(II5ciCvwA(`ftpSA;A5O++_+^yF*jhR8fMp*B(GxW|p)%k{K4o}CNENt05gsJJR!&pCJV'
    '>*;pb)Lw$PaI&$Y$^~OZT%WXcRay>93+y$o7&Uun?>ekOj_q6SHf;B@Sn^Zc!~~UK#TlVOZHiaH25&myWSAylmaoPX*~@2}&D$y|RqR-J6=>fMf@m3X'
    'HnZkgWB~s4TLNaOieVCDowPB^b$Q)jV~7C(38+gJ-?}zqoF;_ENTqGG0>ZWN__o5x6oE;ncAQ}`Of2-*Ly4k^Z*&SR@u&7pgoxi4`ZB326mZi=HtG2I'
    'MyJK_`e$%MSy~28OVXIM(S7+d!1)nsjLwj=?Oy*3vg{#%im(-1-9!0+(K~%$NQe>2Btk6qQ>vU@AEM<30R_B@cFfSvD&2fy1f*V87;n4j7O_rSAg_a%'
    'fWcW@FgRTOA>=7E2|p31M9Q%Cf|XC7{|OF6r56Y%eqJqI@|318&3bpR7!(6uV_%SiF8{m&u^MFAfYBr~5*n*>xi_@v9dmz>-#!or9A}ceGoGrJcSyo^'
    'z}K_wEoznaCLI!lxlE8r`WesfI|G@#42Y93ZQ5?**=uJog=Hg5fYh<IC_$=ntR^G{<KE8}w8YlGfqgY~c=r*p0NG^GV{-?p%G|R#l~(JpHjBC{%LDbz'
    'D5g`8U@_k~qBv<c#s1)JOG&VK+r!6p((Ph|g0>RqsgHao)v;4|#COtjooJO}bX1usu0uT{HaZa9np}Vx$RTwL%W6Tq3?{ZE0VIQ_%jnNA`e<U)O~R#H'
    'Lr35s0tt#u9J_HR(*V{K(}404k;v!R)X%a>jRW7$;H+WT0<T5t9EcQ&`!pQa9yFx;_HXT6qAy6I9Y&S{wM|!ZmRcN*v+%Or^uF*e@=c;hqjf1zr5Udy'
    'fMq4AXB(UgrT@pj1yXbngd-+DLxB41b}@S>xObBJ@*@$bVCR>gXZq<%ng#JB;MVZWvdyvTgjJu%qRVOtpMbHsv$u51BrFggrF?Z`b(zK4cGE5x(4rMA'
    'Mrmv9vTU!>Rig#t+=@(ELwWcS`||~uD&KA+HDaX*pf4yx5-}>6T2)=SgKn;<Lp-8f3Sy>^jyvVUHF4~mU+u}VOc)tn#+|EI;Y~NQWg7rxfF^Ptj)^y8'
    'b*N#z+AR}`sgU_Y64U7CaZUu*AQOU2$TrQGY<Rn$-F0JH5D}+0XIiltl~ogO8OLf8R-HUd=8V8fGRk64v1wzF5`wH_-wo_fOlCE#?Mn=gA~i#w0=(3y'
    'BUEHT7n0MWF{x4pHCC#)f46Z(!Gk>>nO28}hNmRf7CQ}zO<z0n95}^|)mTGe2{)x<G?UgMSr;(6b$mZ!Zq2gW#fCl9IU2K>4f8;g3<~&LCy$?hF1SHr'
    '<Rt5b@pOil1i<=4?8X21w_>xO-L%>Gwi4w)5E)$NeJfuLH@!{`iV!QqO}B2Xe*!n?%7i4S-X*bXyl!^B)JmgpVcBe$b`eRP7mOyzDI-U5?Uw;EX%EUA'
    'Y*=D=UT8zb5Y|0@qNGol!evY965-|8Nq$jY|L>tcxRDnmv9mV+yG1{=O%5>>>?KByE8$}GbgVuLirhh2_`{QZkEr5^tgh2^k~TnD{1_gyrtT1^0^U$H'
    '-lY4a8&CrB2`B?ipGo?_t3ac~PB<B6y1G)iftO;6gOh*~XdWR@z7{l|8iWzZl-N=>^)o3XDdU^y1j8pwgCe<=msMKl&G7wq6iF7G=dne|!PFT1%IoaK'
    '#-J;XDGHfnxxa?v7O;LYHd)=NVt12hbfPgiP5U>}!LeypB-Dv2%UP>Wzsclp(ix3AqCtTj^&ruCD(M!jNm>oP2S8jPr8G7Hkn+p>)%>b3`k+CQy-I>#'
    'O^p|5s^#ap+eu?MRlLJ~Ik4eCo2SOekucrTmiz2<t8Rs-O~$Zt1*}CXU8EmW3deD&56$g-vgvNy={(bE!m(l+rasVL6m_9vNfL?zi#-SUfl!nrwoz0@'
    'noGh+L(V4}?kRA@l`<8^M8Tsn9MBX>ZlKsfPfMi3J%#ucw5yZF+C<d{JG*-Cvo-8)alsc7gC@Cxx@cD`+H)YmfZ!g-g1aDeZ;_v34#9A*IdC^@eMdd3'
    '%b22WVy9AU>xxsgY4L!~vDFk;Z$8+%K*>BjUNyBc-q<F_g&S@uk!lBrUz=UAg>x1vpL2C`kZPy`_thl`O6Gdnx)_XyrYG;DLkj<-P#p<8TNS9DZ^!Dd'
    '7yi#H>?K_{b*<J5f%K_gZul7(3;C(3(+oyia^)XY%pc!S3$l#dCvty^Tr~0_eQk_Nje*rO3Lks`EKzV_CIt=EVlS2G!_Q$i#%PX|2v*Jy+U}^KdxNT|'
    'l9`UR<$>umDZ*p|gV?O*dAV-h#GCE`=4P#|&<93h>rmi|16+cEF(!dVG9AgP*bp8yetB)GUKa^`BuM;<>};FaTjo!=T6}_mYtXsUz@>slX}D$~2heb^'
    'Q_B(X`}Zx?uf|=8I0rliOk~U4QL4)9AaaY$aDFpRZ_7N-SSOx6pfK@qr{(i0B!jCL2<yXnpq1NPO|Q72B=9B$D|aC^8tl+)Wk)g@_*U!Nt={K$E33Ey'
    'Pe^>B^ql?O9)0q}tWaT(@e61gB|b#e!z0rmZpvof^poW$ygV@H<tF6D-opazH+OqXdyc)SiDhHOY(T}kA_uhz)c!PRZ>UBxXsqLghP(Y=_rV^HL9mCo'
    'MXga{XSTsv5soW6k5Mz*3RB-=AlYIbWA?1I=}mQN7qOl78xo*!e~6{it;>n(6Fnmb)YOtY&{y#S0s4jurnhvbwiNU`N1M+?+KblGkpU^ie(Fl~3?&9G'
    'Wt1*ano&wrVUeutku`~oc3q1W$W<a+X`4T<RC>N1f4=uYs*>mF)<+)E8x}V&$SKl@=n@;H=)J?}L@|sojGXW7Y_pv`;w}uV%r@Dln<@{j0J8gTCxM@#'
    '*)H`D;c&><050^R-k9s#$`!F44G~C+R>MY~6C{yBB{__ZPbA&k;=T*n6{|;VS-J9Mq;QEMS%hVhsLeA|-nDO2Pt;T|JyvuW5_Sukz2Lpc$kGO@ViQv0'
    ';h^Qfp}*<mQ5#dLv!5j5-F92}B<hgJ5+a<$YqUdCwv__y4ulcluFH1P=2$9Wsbj1#QlkzYG{YgSPPBSt9Aq@o>fjFOBpedkL~K*K9vuArR$4RRmEf?+'
    '>Jha|)NZV6D~q^c2-KL&g??Ls$|nFHU#}hPvjZ6^rq^-mhHtJCVKqXyJxN$6LP;!)ecaBzu}Or@2;s>$s<cm(?31VjoFz<~LdmAk<v~CD#<U}p?1*<Q'
    'b`4*_E`nL|a77<)2G$1C&rE62HHbQ>v}~|D*mh3l^#OHS5><^1_vFl`x0dE;3x>TzJJ1Yvpb_@Q$^a!Tr4BM6G?TAr5M>NgfI_=b*Hw8XX*teD0Zs)|'
    'O&@J~9op9~=b&qXe9NT<w{2q6&LyrOKbP&AdDLtLclU0S8hGSHDiKlkIhU4+EDm#XxP840=U{Sjoyc`YejCfggmN4ws7<0a88x;j+<(Y3f%;@Nk=cxF'
    '+O6qMTFqgnoT7D!)@8KR?Us@}PSE*8=QFzfDrylY=&D3lWpv^>Nn|c3s}ot>`^1>$_V(l(n?%}Vq@nf!X*(d|b*IOy(ng}6GDj0oNOmWeMWwBzdRqxj'
    ')rOA43A(h0=qJ2p3L@kN$48eyBD47%skD0n&metLp5CtKpl0O+Y2PS2LXc9PhF*>wBy9xM3norE^tkWu(Ihl--c;HTD$Bs%h6`_-!HiL-Z9zYUI)t}8'
    'L!EYp%6CTigwrdl)5_|8!0Xxf>9nxiuy8QnN6{xaZCYGy^3*<wqYhEJfZ619X^pi#`OY;eLhrfqL9|ISO=1t)y^Op!QE7ANSOp@CGp;Y0%CKJ?w#x-_'
    'o0F%_p{qVQ_88dgiYk9PZ3bQUpH44{k|w>Y=#4e)2Hl(@4|G1GGl@Hx@)^2(xF&5P6Wdee)y#xJ9SqUtSDQ&Y$>i5eiX96%r@d`!(jGEqlqw`fwSBU*'
    'g^V2~%PzLj_ohu`+&DnHh^mJi!ZDz<os91T@fqkpopz6@#-0vN+#XHZL-bP?B~<W<YtdVbv~f&xh9*B(Hf<VPjx9bA?SKFL%hNBPzUO3V&sf<ra1`RB'
    '{Gxo9*|byWCsUX{MJ#2)kELB=O{9?P`RHZ&OxUzj=;!$G@)K^;)}Wu`6up^Cdw_oO%<=vi66)To(stm6nqxDe^OJ4UX0Y|C!PPlI+6h*6!WvWlhzATP'
    'b2jbLesHoM?)Nj=pMxRhWNAM**$?QeU*Y97q|BGfrG4RKUsw=xfnRmUr1=?fX@fX3Cc>Q@K)xI<?G^e7v?x~7N+vcv{Px}_q`l+Zn5-RG?+_#3>@Mvl'
    '=ks$I*jfJw$F=rJQIMb08G1~xD2oK6R(g>W=tQ7}Af%y8;63@XHa*L?_6jtXws_pbjgKBEenU>-69tQ5dkmntYi}tDu`#`&H0fbcpwq+3m<47rWA`VO'
    'kp(%}_tfW%OqWb%#k$GQrkUar4QSFOVPS_^8L9*_?AoOdu?!<Wo90WU`3l4-{`!482W7S}{y<iZ93V}J^7(rgvxoh5TR4Ypcn?5n^RLut?m0mPjm)zL'
    'RM|W36$mS94|V>DK~RyO8}v2LsHp8^zoIx+qZhCW2VcU6ayA*pm25K*lRnt6c!&w8d#d?1na@!754#1laC4~r!RH$?-0t#|$k+SZ+X7h7u6>lPRv2?|'
    '9{$RS(?_!(lO`L~WR!zDv;-D~hh0%=^pzV|TrB3R8Mes1*>BP7v$)LBTtagrXco{cpm{>m%`<3e9qH!~tpT(K_E2SD1HoU<prxgvpRy}f25=1EID2YT'
    '!LtLR9To%y>AMG9C*ZmV90v$1h66OO0Ir-=VDt17?Hs5HCbw{4HnFou+{#rW)_w0&g=P$n?)F7#*hC+stkpmN7xtreR0VmeV`#3cV{Mn(b2-%^$c56&'
    '*a`$e{3uA3K<d*VA=-p%wT!_e8&l;@eGnM}a3u8@L^M+w!yd))k+Ab3c1zZ71d01WakFL4mq5G!WVX7)!ac>!Y|4od=xFR#xMD&WTmE9bqSJ(v4ol7k'
    'D&~Qc+5B!s-my9wh<c>~VpPi{h8fZGIRH0)5Cr=yLe?pQAwmwcK)B~1@}y~Rl7pt!Y?9kx2ns0NsI1@3S*1M$&g9u&$Ip8XGhu&0jO2{F-Pcl#l3fn$'
    'XGqRN+@WZ^wCD&AMiYPF8-ij>D)*w;5~3mMD!K|y>4NQHOYWb$ZjJXuY|(%vc-e63fNw@xm1vu@w9$}ZoEUp5*oMkw!QV|*qj)@GLc$~tO8;5BeW#;j'
    'TyeZC?MGv|W6x}bkw&+R*?a<nx!G^p4=W5BQIW(}iKH4K`P=<``|dHL2+>_w`<)x7Zm^UMj;q^DF;o|5F>-*pkr369ZB!hXGET&?8@{+8e=dw8@Xjj4'
    'ccNR2J5=y?m~Lk4St<d7|A^`CaVr5?tRd1=<WSAYVnvA~A}O{A?{WhoD)lx(qW^!=-fTIJB+C-~6`K~B7O*D)m~OT_m6$I8oZ(x!!x;djUpNt*te6Z2'
    'Ll|TUrv73+Vg8_6=4~FUer)cd_o|w?xx>jWiOOUK{T@?O)8($`CXsp|5?R-dWqgtuj*aOujzu)Jze(y>6ya2X`F;hPifYA%(v?N_{HC-)TBQ5HRneMS'
    '6jfxv!6r!aB|6Khiu*8Jp^6A=PCiK;8p+LQL#!lS@<+u@8u@J3c~%ZZNd32)xWE&TO+fw?ApO`U!CV4!uYd_}%SkMsu<R>XJpU(=LP9dHAmvRfMzC56'
    'f;8BOFy@X_o^0A~dDme*$o-*vUSZ+XTGqHPzrhxE4SY!CWe7)5)+eZT94h664baMj1{ZtT!d5ea7H9({(eD008)!2D9T-j{8DLDAS;o||m++6jzbO!D'
    'h}}r$4PsYg*fgV$NFk-52v?wN4-LMCV<m!;?8I%(I{~XxkelhnHn9BTu<(Dq8DdnDK+ziR45=8s6dcUSd8sL3PF}KkQy50*VcG!Ctn@88)l6z6C$w%r'
    '7oi(syV;`qv=l_R5x_itQy|0uv*q67#J|r)F!!Ky;u2hK;Cd=%M8v!=3jArxYK-!(Y>cRW@SdF3+mpZB9Uhtff=tN@2l8Tmg3pMueK6#funEi^fx%wj'
    'dNjs*F}NI1tcT-E<Mgcd73GJ)qKae~3yOQ~FuNlAfT4U-pxh8B)=lt?HyIFJf3tZ6BP^V76j2N~94rL_;x4MIs+L;y1tuIF=0x|#bA+cX8lr1Sq$F}('
    'e8Vfu341euHL-~W!}<)C+ir9MOI5d{*gxGPcF3_o%W@@>qdEBwHl|3xzFZE<H}%Zr<bLePGX#nZ^bF)r0mZ(PxqKaKsDuz0Uozf?upI6<bc2m&DM@QP'
    'Uu-uq?^M5OwkKiZGyw@91<i4{ne7J834DXYa*A%_rv>PSm{M6_+lFV{NpY{knI^j_o394jf!46fT76B+iVV(6N`is^9sg*XUp~F~PhSPi!OEN5Y%@gA'
    'idu=>5hA$?w&WceC`RkW6j$RTOh2eaOm%>>po7H($b7VLs5r|F&atm=&2%bpQBUAwE|AXxGbkUs>6F&>4W*)eD_q?K@tsDFeN@}AUY8Q_G2!kpwon;U'
    '`ZmYvTO((rbAc0~&Y?&K@3)&VBsD<mSQu~yNfs`F4Y8V|1k84%g>p;>9?-5%^Bl~kYDFh@+;>wGd6V+*$`@ZwKdD5Ux)|7T2_IfUn1YQ1m)Bvvm~}B('
    'wW-CDD%l8o*X=iCT+E`wb2nd7nt-T&k%kVFe68gx3d!acqyljP6u#<V&Pp^DSi_Wk`118eyl`zPt#pu50{{3(;X|r_(u{vF_hCU|N)y!z+`Lg?;e>T;'
    'Ka2|_&rJ2Va3pRUtNiCRg+BT`rla&$Ig6>>W~~cymcr-9KXOsIX2wo$@9@TvIZ(MOXq9JKRc5@(o!Jfws&Ie7V)p_)4R2sDj|;((iEx%%BVvXq7CX&n'
    'UDjf=twcq~-}%As_!M0s1Gun0w2c;oSeE5Pz{$<z1{HY?$#F+8bcN{Zz%=_L>o1V?Nw!rtfMq^8)TlM&A6Hk(&W*>)WaXMVYw0suMhVB@-3}RQ7wg&L'
    '-C}lt@6*O&5~=_pR#%+Q!1?(5`R*Mi4O0~JuzuV<vmXh|mf_r_izIGm=U$v+Vu=A>OL#GA{3@OVS#!Xqk4OgsjL=<%$TN0wd_k13ik*TNAoi9^hMWV('
    'ZnH4Va2Y7A^|4VJSIXrJ3R@Kw&6qnDi9ICgi$GS_=tZsYczmON20{|QT;o0W;+Q{T`II!?(zI%a8Em>)NfX0=OmySIu~%_5Iub&IQX&irTe>K0D5#>M'
    'n5-uSNey6(jpG2t!Z4xrbQ3K%u(8RxhYg~wEB?2*M>4n=@Jw(a5%9Y(LcGD5ff5q2{)U<vJQdVxl_;JqXkBmms8vQ`$BNwfeviZxYLN_fJ-H_F2<=)o'
    'gl?sFql$!7I*kg;zBp6RuY2NPiAYj=6$v4t2}5k&nEinjOa-bR^y9vC8)>HT;oRYa+Wj!IHXRYvgcTE#7I+TR2SuqK46&<8H4<B=TrY=>V30*DBlI&i'
    'u@sns2=w}6a#m{tyld%?-F&x!wS7mda0lkW!I`Ds_Jk|AJS%-6y2y7Op0Gs3Crl@J{%+EwqEH_tB!?5cPLrdo5A*lMZPC5mux#)qm-Q|hTqNs4?+n4`'
    '^VR@S`eiN(_I9mmi{YZQ!P>2K-2wppz7J5B+PKECiG;25IInv$r)>RfsiJPW?HCH{4f;@4oih~^jULlAViYmZ*+39>@H)~Lj3L4ah$Qd&L$KaxJH*hm'
    'Zs4lE+s#oSgGIL3&6D@pR9wGR76dRvbag`Xmm%6RN9>(m!B<gR#2I&GHTsUQ<jHoFk_>!X6zm}dY^JBYQnh49VJX(*e7(o>oRS3SheOG_O_JuE<lvdH'
    '_!rC6Oi5SY@#u*?8Eg3VnXGl*25A}y)8NCXh5>W9M?wU8eFUMq&9s=3O$Y<XR{Sz;+zP`M1)v<&c`-XvMA66Dll=8|2pNoyGH`EByc*}k-AX$ey|LV`'
    ';~qV6k2FU3Mg$mqAJ5gP`nsYJE4=Td!*quE9h`IXW0ni5LVAMZ<{*9Hq~WTf`mB+4Rh1Ium_I!%CcGpum6e+}wGD9=fT~jVU>v=__{ys2T7PB!BAY-y'
    'fwK%V(7$pC^OA28W0N`NUBa|JbxbHP@!`C1r(#8@wKt|KX`96Ips^g~CtYAhTa{wzn7i1^Gk7}O-=zlaIjm<>Vbio*FDBhme(c2cP{Y&l5|KwGpK5wg'
    'kM7{bum=W)Y7<(mXjC!xI+|4fXR4d@dIuQK-C2%d22+qx5#YEIk^oumWpW`<^ikzQ3QZbPc{!50&HBR<jRw)&cc38;VIQfsibRg=3y~!1(F0QlXdN%1'
    'i{~AkHAN3fD#Z}0`Hq>F<&c9_+1!7}R}@39hgF-j-rI{-sK7ko3*PT=!B@AsX?;utBfo^JPWb+xH}`PVCZ`*19)ZA)=Gt?IyL@Tx+v5HpCo7(|cgw?q'
    '&O)0tOTuNpFRlPa&sw<mO6@hqyK)qAs*sXJi~Z|VN2*Fws&9;)6-;rxpwul?x~b_V7x6NW?qa{rvR8ECE4Vjgb;147^q~0Vaunv(UeDiZto~Ud!x53Q'
    '&`L?sdIv~csAl|&agqsxngO@Q+fPL>Nmg2zmtv;q4`qqXj?4C`8r`OMF+qJL-2QSa6GBCS5~YzTxV8WIyGn}hsj3?#j2pO+xP{s)ZX}B0QmMKziO;w>'
    'P49;+5a>L{fFKnET^Fz$=7yS;x>Cq1GhElGQc|1Ks#6m+FJd1=9%k$k2-@g=QI@KS3TTs>)H1j8cDGpx%ldH0ri4k{;9S69_t;d5S(i<vrmtyn?=+QW'
    'Bm`u3px|ZKuStk~RzPQ(%{mk)uKxAIm)}1BRs8b#_s?HG{w0uBG91yrY)Y#g&Uf7wHwv>z)}KDTIIo^4E9r?SRJQd{hz@QQO3@FJ#3orGYnR5p)w~_1'
    '+U%N>o0BCyz#PsKHxG^<?7LDQK<y-)Izm8|tA`mT(}%@=%NBXoXH5IYBMMS+AWYP>s-|t5AxvsV#aT~Jz7Q6VW+F%iB+(?rol|joBED<?MVVEoE&lxN'
    '<G+1K--cu>ql0Lz)@U3v3`NFQ?Ez;M4wu0d9Pr)t5vgmJi_M$qW_Cm1{YP1;564HDS01I4cHy)7k5cj(e3X6VQNn?C**nd4`Y4PxIgVSA`EO|5;UsNN'
    'H4aQui0k%`Uj+-(f3z+WgIA6sW3U>xbK;p~r{8{i$U3iN)Cb7Xv|8D4t?oADdIC%nX<GQ0u%oQ<BmVxxGq3#eecWJ`6cGQM^l~_S9x!R}=Q$=h9AGz~'
    '$_6^X{SHIo8w@_aqxq(e%vG&7EhA$eo2=WRmJKn(&m3v1u5nQI8s_Qt&RRTl!8!^9m|$#bifi0aQ(XKGPfUM5i>;V=)A!b%!FFNQ5p+l$Q_Aq8{dP4z'
    'kpbtuf$JY_@y!-l586vPD;<iG6%<=-#F8wbXyEn;)HB;4O?8>QJ{~`~=#`5ovaF0{*wFxQM~_*qWRR21eY>Hvt!UKSDUY&B5@guZ_k)kpZI?YPe-s$I'
    '?5hvVaK8WewYdNN^Z)+2c=zF#4@F&U>10q53MJv(=t?CwD1{D@kc>S{aA9H~lGIl8%C45;9l*x2+4j9Ip0Tq6HMB53`3o$*I5MCqkuUlIgjVi;4wn@X'
    'zxF!C&;<#I)|=)UM5*gl8`5WFGKeauE6Y?|CBjia82&r78`hxb%;w0ElvQ8RwZH99I)`({HZGN>w_4Dz1QU_F$%J1%J>h=<2+Cv!mM=X{g7=!0XKRhF'
    'VhXS}j3j^P9)5Wpk}#Y?VA2?EXV2Zze_F`FkNLwH8zNI6$591XO7r)fvU6mBlIDYe_9DQ4J}l7lJBA5sx(+2XyVAt7o1(X<*M}8~M`D0no83^>^}m<6'
    'eqsn)_e~uJ53$0lN<E$~f5<jQk1u<SHBNnzF7~bdo_DdAJR}&WE0tNfzV9L7l84kLO&GJZe&0jtiyjg?4f8|74}6u29@E$~VZLmC;4#f5k7?De!G2ff'
    'yQi$Z=pj~o_}IC%{DF6A6|jY?5H3HffUGC-sm&}FZ0p`~*f>#I^tCkrvfTMkRKMu1=yh{QT}4E}EdT!H%>cuxX;^b=)d`GTA3lb$PL33WsgGgwYCMI*'
    'R3MnICwUpHi##rfsx7H>`we`r%`NJmSU2@<v0}9jn?7z4pqLij5H%!FuD_aKUvs>hsWcwJt|Rg|YAB{dP^Tvk=*yzxzX!W_F_(Om4s+ZD!?`02Bf$bR'
    '+w<M1dMmni4)Yy_rXP3R^l?E=aS+gGI784H-zc2PC@*)5d+bl5uge|7XG3X?>O|k~K$<-`vqgsiUtyjfqF4=~!Bn?LAlvR4*LtpvkD<7>Nn_!9_#KfY'
    '_Fe&Y%mLLe3^u`hfY8wd`{-|EpjJawk%Ig=O>|mdFF3m(c)X5+O`pqGInzV2h9%P#P5DxxzeEhAXIWFc1OhJ3BTQ{pSi0y&68eb4fBxTi#lyhR`GINQ'
    'FV>!eKc0Z#rR^%hw^|akU0^wEvEO11(9f)3O{eqyhPqA=9*7>nfku?3fGfSC(@XWNY^P3l^KLZ&Xl?qkWUg}08|@Pu`(8-rHF;GY5=ak+gzGFDQq$9r'
    'bh)HKRZ%M%g0xma8V-mpQ_q?{#7Jw71Z0&B6dhVU0+NKFU4@GZeS!3DD0=*Iu1uegv$RLxRunlx&iuc5tcH6-aM&k!0B5oThv~H!>$@eH2WX5HlilGh'
    'bqv4n9#d-xbJL_z0>lVx+GRYE3VVkeZNI{zxy}DlWRNz8td1eAOBVsz<HwM9%Q+_QJR+m=4x0xS8K!k3nBFDKI>+QXgK2l-ovj~$)v#5Sy*386rs%rh'
    '>h#UDBMPu(hsnKkh|kxkCa_tC>B-A?iJNxX&E{+PDQs}gT^c;9QKKS;n7;o}T!1uuP66QQS=gfmW+w5)({6(4U6GlGXaD^|IyZ1E1nl4C;du^wYRSz3'
    ';E>I(bW1rEwdSX^WC-Z~Os%qyJwU23P^o0m_+x7(;v20Qz(naE#INep_^Phe%zyqP_-%6%Ui(h~y!+AVX~u?6ZCm(KiTD{!@(#l^!rOk*)tZRg2+>!R'
    '+An8i<LQs@e*5E>fB8_{ef<5me=DZnKK=Ro#}EJcKR*9etUiAI0RQg8uOEtEKY#h~>*oPd@~!}qyAbKm-#>i)*u$y~5y%Qw#SKQYc_d32TTF6$+%Ps-'
    'xVb<zZ&KeJmo;ODi-eld?+3XmgWL=3pW`j%z$r5ORW-c{o4uI)nrDjK*VJ0{q~wQ?VdWv|4sFu|WkpLu!wvSmKc_jALwHNn$xWlq2MSLA^jmaG{Z%}E'
    '{Pdsy&)466`B2<{`}pO{$1fj>pNc;}|5-FPy0I0^#<A)ZTeo&(u+g=0$yj@80KwOFrH8vW8I3F5X<1sl689b6Mo17HK@<g+5ow~6Vxf1?8O6epzSAgS'
    'BIfGlZ=IBY{bD!A8f(t#@RrYJcvq0VE_)Z5!|H5I1rocrq03P^K#4YH>$2dI3HRdSMK=2xzHxo0P*qGpmhr)B=azcp$i^ygn*HsZ_+v3Yt@!hozltYb'
    'm}N5nw55#A_RF8=tH}maX6W9diU_b{GNf$FRI!e+)}fI$@UF*WCF_G;GsTisS}#o-X?<BbotzAFghPXA3`|xZ5go2tGe8dg=&y_oxww>^`*d?>Ms2`&'
    '|9ow*C07qGsOb$O42d#e|1jSz?(PtFhBrDH?*q};_KHQT{?g<?g((AezQqlbJQ|UMCS}iBE><T!2y@fs$c#b`U0ZNk^kV%?WGR`VOCYst)BM&=)+uVT'
    'TVQ^Aa)I~t<v7DB`!WJH$x62z_dQt+5`7}PZm`~F0F*}N2{WDqd7g!SbOacVq;_c}muoD1d|h))*<|_zJ*0L=3#u(8bcekGv(BD<mq~UP^BcFrH*QNc'
    'Isl03aTZjEEUERe!(dvP0yKE@%%Ntv&3Ihf46r)9ofB<D_l}ArB8p}P2ugASU^sMq{-hl>d&#V6i~fWG$nk__4i3xgNGA_dHpqu(Y0uDL%(l_E6nDhZ'
    '%@C8nxl0~c$EaW1_P<g7{?W~k7e<xRH*KV1s!Uo+bW($bjhWF!N6o-{o!hnB#1lDqP<S{6nDnb_#%K0?r1+R@4%6j4|I$3)>j6r3v|8p()y;dDVvgdx'
    'xGg3fEJjKqrw>hOfIS#*v^gZe<elf>-6oh9DNW-BM*!n_bA+c6A<GT#2Y9>;qdQYT>ft1{zEWkY<#nRpl$&c(c25kVWM4N!q5d0XI^pQs$!Mm`?!!ws'
    '&P|J2Ngu$o0mZLF!Qj!?Nuj~?LOBvOQ8qRq{m3%Pwy5PpQome97k0cI2OlcAU*fgce@lkazNP@>M=sg{rh;Zy?EtfC>$I?(?@c|2yL8gH9U-4YAa(@R'
    ')as^%ZxAj$%z2Gn^CuoZ%?A%^eL93I&Sa5qB)M_>$iY!yCpA2heA|O@oCF<G|H-@^!I>15+>jv@I6cz@QnY>Kt>Zp=<qXbNPDPq$NUwP+?M=E5GVpWU'
    'DFh1g*JY-wwVdFb_I<@!T5GI~j9$v9NX@2YJ{pvCHLKTD)<gct`JOW$_w%Rc{%^gl(xF$E!^>?q$Cr?~fdw_2dRq-4lD5pAP-(_GPmO%j@q`FXAED+r'
    'RYVxMswF1@rE72c(oLh$8dy8|xt;i0QPW@SOIV=kE~>F7G+4JO?&L$x2Q8Twi6WV!?8amlTEm5fn%55IQKnpx{Jhs0h2s?1Ot%MNYnKpWexLAl4}RGO'
    '{b^lFh$a4%G*nuwR|tR!3F!9(nS4(%nU=MR!#);D$LiMiWSGbtUKcSX<r-TQ?$H)=XrIu!1&kqRo6cmqZ_s~@RaeFQezAr7q!2Zjk0%fqAUIVHU|2q_'
    'rrDyKVQ*Tj0iVC?Vy;y5fuj{UfKfHlurh;h7eC9hzRGU1>U53T-bP0+P7kuF+A2Y64yLm6J|33KK~45E9_cv>+&fGR#-A_paP1~BcGFS5apWN534-6}'
    'S9!!ptA#tlOcwXp&`t=+R`VxVaO;gg)B^}thQvT-YE8yVOv6}j#5WU{3S9BjUGs{Hv5>0CPeQ`2nyO5=E)8w9!N}i3PeGt}>j8vPN?jO0J3!LpX6!>G'
    'k5w&Yvg7E`G*)}El~t9RLU6h}-?7O*0q2RKFlJViD?q^{Ck0(U(G?*kvYZd_X6wftZEb`LJ(m$fwJC~N_f)v&BI*SX@X66F{fK)HP}=$k#g^*)XT5SH'
    'Ou@@&G2dZ%dUD$b@peoR<p|jsO@N58${1jd0md_KI6wsHUQ{qbs+X*c+0`*{H$BXETkIhO>f&Oy*c`jGV7w)_BELaP<taX;DUJu8+Q*lCxzl%!U+9%#'
    '+v5-ATvMe{hZJJ%X_ANiLqQSEh=gJzP;GL()zxXwo+twN*7%vSQOAM0_86TLNaBmT5%g@tedNKR6ya<(Sf<FsIDLv0#(2{4;Jh6fTnavNU$KZ5^8g~2'
    '0F10h?w%C^^2O5!dgTL$as)ZcPv8at$ynZY6nmFfHa8A4;<AS@Fkp<}hBSeVI}TWP)C6-f$98W`?#Pv~_5z5lA@xr^Yn}^lBhvA=7<>D2=)Go8<Lfj5'
    'b=vB98ojp9Jt5{>o!Z_d7dD43WSmR5rq@K4Pl+Y3jn@wR%?xh5Pw4i$$7@H%n+0aM1D&j{xOMW_LoozTSdPDu4v#y_eg8-5uO`PM3<=d|BV56$30Kee'
    'j4-6+A?_?VJSpNxG2v(3;f-Iz`jlm9hU04}XsqcOhLci61%>xs?8v8Z9eMt)Rg)C85pFlr^Z{evOTP5IIjhq?M=kCRZpFIl0|l+6#vp*?YD;6`2h16w'
    'Ng~~ho6ggSkkvE;GPyh0a36C7iVh_WyUwQYyIi*dS-4=!OvUq!%C12#JP^lrNC0q?RJm?Wre9O5?qTY1BK}Esr1Xlf+MzLSdXUE0MonlAXRG?KC8DIw'
    'e2D@6GLP^#0}Pj}dG24M7&w>R`Y;D`t*|#fN6CEJjN%&&hOo0&h>mm|n4z8PT1*?+5C0*~aN@u;OI}v)I}<Z|6q>JZ@&%4$wh44N!ga|^=78~!;krJq'
    'tuzzK#ug#NxnQ{K4D1TqWzRQnA}C-CK<nhk;tn4dfJUsk%(1Y}O1;aLui})$h$X2H^QZ29v(Adxb9syM$#c_`r{#r2AQwk+4xc=`buSZBaxk-)N-U2S'
    'hzh?O=<a&9G~5F^i`O2<)F;1x_~Vz)#s5`IKY#iB$A_YJgTQUYiTYkWQVqYN9CJNQo>R!8zZXxMJL((KgECIFotD{rmJ7fF6+W-o0ieaI!!(wCU)-jM'
    'Y|Dq1rtdv-P9YNd{_|lm>jr|g#+F)7oQoRe)W-^$k-E<N5U#G1_oR_?yEg+aJC|l%l~QX2AvKC-e(}}jakF1gcIT<x#1rY1c*R!S)xs#<R?FcoDpE*q'
    '!}0Ezsk5a<1yMI10I8DL@>`L!@)2zE#Z>Gp)@kDWn42nm>SG|^P#O#q5U(-Ea<Y@f-v7FpQXUaf5@gd5v*8SHu9*}!W9`{&*G&)0ZXoE@oNFEVEK(py'
    'JOtP+8zsry;zOV4==w;)VUI{m3pJK$CvpE&tTg&Ql~ZG_HL*F@R+13X@hM-X<Y5gkyjs#N|8dNIPEOZ>bgkj<mdia>M4`K%@1^Zq+%RZ>_GRl*57mZC'
    'St-+fi7rMwcrvOkqSLv;cuGNM^h1ZznVJw0+cq5`ltU%^kB!bf=jFttUQ9r(`T$L-2S?NN$~**s=Jq^9TI!rd7sKrs^Ck{X)aCNYP4b*_p)r`26C3E8'
    '#`$>llUknGG!10uWD_fk<CU4U3DQw7C-ngVWUI%L=Bh>x2Cmnl7Y2=EGwJ#;wK5-oE=hyTMiw=b(V-<6Poq(F5C&799L*Jt=a#kGTVB`lmVlADuCe5T'
    '`RtT8C>F*u-#1r?#rqJx8GJ}$-!5$XV(UVewo{$}`iet_f$8nLR7oY|0<MXK+8QiHwy_g6fT_7=ehn~);<aIj*OVvlW{dS^-yN`@$Z_PS7J4>>Oz1JA'
    '$0v02I(pDK@@bB$qYkr<TGUo*Iq59<=)vIHK&uIcnX=Ll*(&!@!6(DTvVqHu2M}Qz8|%P*%Vt2vX&R4yVKk+0t+b|5zHOzt+qNw^il_749K$j5j1PpX'
    '4R8i1)PWp41+$&N2^5i@I2c8(CrC#*YCVy&x<=O)i!s29O#pO!zS_HjWk@<>v@m{TrE;%mnSIQ22Cc7dl(*BGm|XX8*imO|nIWU>K%TwT>3*XgzL^(|'
    '^-XQUXDn041_5ds;mqL266M2`GP_7#8#%^pY?w+EC0uOMjHHNoIg4sb1ct~rYYJm<bz2o}lBhc?>I&ofqHPlOXGMKsd{0y;?@_cmC+Z4%W~G8Q3D)NX'
    'ec}6pX^QOGDQb!a1$A!QG(q<41lhvoPppM|6b0LUrnm55qK=RB)vzG{<NurLTCz%^1cnR6gJXBsEP|-o(SEbU+JqX}PrMWX+uA6tmdX0oZ@YEm^j@sk'
    '^<)0=ciN#*uF=!8Y#e=HX1DF<<QASu^gxw%xA)Fm!)}`Th~bn-v~-qP=(yJnM=gW?OnG#jrF)GCDc>vIXmn0sG2x4B4wr3TZSTF`?C@PWd?)Njt|T<N'
    'kFl0pp2bl?Zhy4v;sVv4n}!rvBVt0c4r2WA2V%l9MM7Ssq;*uylx9OIiPy7OrDDqZ-Et?gfmmX9Be`1T>h!p=K`M<ripfC;Vr_c2c0_F7Emxbh%IV{T'
    '-xW*8#@<e-bEBXHKb;kNbsT4XBxo}xq~15MX*Qd3M2vdOL?N9^o93i(BSMihou8CO*d!@GBE>V=-5-*y6Iqw0#g7g1#7B2&Qkt<zIrY(n*9GO;lS4iF'
    '=svw+dps|$N5Y1Zx~6@b1pSEM=vR;?R<NQ&J0c`3bSTt_khTc}i*rn<Ptbg{%7k{a^lg(`uI=V`%lQ*l4%~!K6cH>weN~aP%SZNW$l1~8St?g0IX5CV'
    '{(5QKv5#dSbZ(8+X^oBSpm5?is&$E(k+UyE)97er)_Q-rk<AcJegPp8EMvVO`aAIWm%e?8;2c$-t0{9z6@PvE^7Xf0KMR48tK^(<F;^Cz{LjR}AeK`6'
    'VQm^$Va!C73b4rOwvRa)pxI9r)F%snlZ5Juz4MW@t>&JWY{a3Tf3fBpOSwh$ZjaJPSm{-|V)xGD;;wtog4FGv@^pG)0j5SD%C$6cq-p~0I_z#hBi5}e'
    '!b<Vi!aZn4gy&CKXu6MSl>7MN)c_{Q;x#oB7z)47qMH!m^ez8!!}4J$UkQ&W%Ab6PIPKNoC4!MY5_Q!CKO((a%vb2&Lh)csPrQ;=$n@|?cC(vwNF^tM'
    'x0s{};`r;Hc#E0N63A-FRjwHY63`NFWv`^0W2`8?d>kP1XN-u=C}MP%jYArH45?Z20NI{{jPKkBbP6cSSmf|yMGz)?w?sEgYQ_@Z*@5O{vNB1gNM)B5'
    '8=+~NfMv5;N@h>|`HKlJbi{d1UsY50b_Lq*YI4}!BW2(15?dH;3K7VJU-;T!`T8<>?dwt~5spVj#<EJRTi(`mNrSZuHa2EeftpDNE$6zdr<CeyBV`5V'
    '=iW?io^Aw=OF_u{!n^%EA!f$T#ssi}gUm-h_YIf99C3y07x!yQaRb%%9(zh_;^Q%=M&k$2R_hE(4%)r~opPklfZGsC4JLYLB^6G10v;FW@4UkGi5y;W'
    'M@#gU{7<V@2-l^F{C3`n1yCay2ve(XqkVH*@O{H=K_9V-GRuHfhDpyQD0EFRS$2m-w2=sU+T97k=IwL`Qve^_l+-c>KS>!IrwCNCg#+n0aGg$xLhI-G'
    'Ggfp)KPH@Q2fR^a$(iz}`x{wwc%ow&G2189rm{4N<d8iixNHuz0=cTJU+f@Aa-Vt2`RS|AzyJEN=zjn9>9;?A`;1|w5C8W0k75P?D~dFM{y#)jo2rZ>'
    'bIEQ!r`arG>zv4kl2M>kSYeE`?{5ZZu`tV0rhJ>BMW~#ot~x{%sZ;h(yqk@oj8oMk_S4Pd^r4`?$3{(S<VwJaiDEC<iL0E5EgJhD_k3oFuyk^m!PzfD'
    'j1Z7Ql(Kr;6oZjO9ds={bU$|s+^U>=r!Nu3DXk~b?2s8QUEkDWmi+v3Tq2o~#7u&%t7A4$idIq;SoScdMZZ;`2%6-Tu3$Y+bHg9ZS~;Jha9yS(5B7UA'
    'M<EwuB&$mMX<-JBTne64m^oLWq0<>QAbi45kXJ2o(!D8G+t|Ss!>OwFHr1I1z?h!*Zlj)*@WeZn{C4EXeBEys)bI~aE=kkDd6C<Y#2DZOeSF#%sqwm='
    'xC0`3b#7;K<VM>)r!~g@z$+1dlY<E*PZgYIhw|?eWI_^*kYhAyJmJjx*7hbiiBHj#sF0N{I-{*LFCQ#z&KPHI{cbB<=Ac5sel)3<V<l~8QgcE|ABK6w'
    '67`B($f<BuC0TCzPHo+$n<*|(m_xV;;UzHht$cCcUMlSVg`RaqXk&$KO*gbW9Zu~<vELka)R%hCo@0HkiM>iufev+IphFIk8@S4U!W#mQ7I<ib9_JgR'
    ')o*;9?@G-zAJ}Cw!h8>#gxq;z@m;O1k<m{e(7H@10Iob!o<YkPHu+D*&+sm(#?ZXjPhYW+-B#BCcWnhaDBCt$=4VF7owMT53c$v$;ljXPahDR>pFGV2'
    'q<s#xU8-jyAjc?SMj)R|)wVse(I0L&N+>(Ms+8ZC^?arIzFR)8QAfYKD-PJ;XujD27BS+KB3pZH<wfgqlrfrbn?Rj#qF^f~Md|seCtmxwUBY%{uq)Hx'
    '?b$Mjc+Ew49E{(t>@i%;$~T&A4Yw8@gKOlYsLw&9>o%~_$Yd*cSipt?dl)wv_d41?Zgi!xK$8UOyb)1;QrbqSUrB}?o+LvDL|>+GA^lp-?WzsU_<}8('
    'LRPGBq1L6k@I6xtot+50?$}p1T%NTtb%sP4lW?<k)F)Tk0pBq}8}i}8yJG*N#U=$#Y9qb3;3drIvE-k4WqKv!wM}s!jy7_5?(HO(%JMVB{JV11FQ49|'
    'ZChL0>f11b#n9K>Yvn9nm*^0zgLw9^5G#zdCV$t#?BJmeyW^K%DD3p@kKdy1TgMKID9(LvVu?b`ssg>4o_a@(?UTOuHugR$Z@ti^)wQ;q<j8|=+Kl9X'
    '(oLRW9^qibb0D^*x73le#$r#6ftlJ-+AeD}vD@VKCU!B)+SEpbCL^k_%mfG);{+?K@iv^_afUX22&A%49a%(wt99SEO-nh2m;{A~4{E9R8w%RTZ-8^K'
    '-e_#Get~~-t-u_s7!qhmm;B>&1$zzDgD4)~gpyfB{|}Tt04Bpd*=-Vn+N+-f@5)M+Qd6~Z19;|C16>Ksm-odCDf-bzg+?#zio?@9Yhd{p@1h3&FC}+Z'
    '4Xwu64x1-4q4J(er>Jou*VFB(!@(TSlb6(ArFmhm2uEuDMH(N;khHYW#J8wuapziX@>7HPfYGrVnhF)IWM4B~Vo#Kuq@+h@l4&P`bleM}<WB9mG~4r|'
    '57h?mzBwANy5r-IU9CMjEh`fD3rt3bmHBv>P%&$Ee#F_qB?ObcStgK3e}bGmbV(s692v_DC(H!!XO&S>pq3AcqX+Z1B#(X+s<Z;7u@u|@gK;kn*!Q#<'
    ';xB7k>D<*)3J52_^rvT{Q2X=aFaPrK%ZI=H^5M(Jukc^re)*T*e*gW$n_@nFcz!P$VoT}u;{UA!y(OJJGvYrcs~pbx_v?);MVTWwDobuwbOji1iai{='
    'k2$JIsP=VKU1xX%pM$jt*1Zl(R~p7?&Oy2a>0gHwl@((o=TKAol*U{#(6<~;KL;7oK&#gwV?!y5A)G_4(m*emQHnnSW_1SBDgQoTHm_rn1%@4W)rf|D'
    'PYVp-rST&Jn*`i`53qJbjnL*8+A1`aBwq|2)uuA$CJazrLN(ums<AKzcE{sQ3#_%@H{Q28!~PgHBv|)7u+qPbfx|IyT4$~Q9$=1mGHi7OYto9KOufrj'
    '1YA$I+}P%QS{6s^qDgDHEk~9dH;a|Cp|_5^ct~1W$4w7Z7h7x0rJ#D4jYFnmtgbJG%n!41RG$XwE+1$-%*G)@g7lX{_71Z#&`RIFDK7=h&ayGcHjVS*'
    '3C(nN0B%Qt(~{*TsTTof%avDwF(cS*c9}OLlbg4thiBYlah8is>ZRNGeBWEX@PxMag!K8W`{5_#-}C4RReHjYzeD=CeSBlGIciU>>8y<6ctabU6=Qms'
    'N!@Avg_wFL(S0<En)Y;CU5MJ>(qo`Lfz}s-W;=QeF)1c(b0OkDI~ibFtvtZyAmKdbx<kj;ERl{BA%zyN0IP@FCfw>m+(cY<4p|M2z{w7mASdpd^U$do'
    'tGgT;8JfV<=s_n?pIvPJ)C|jhI=>%Ki?IGpj%}qypC_B*JXm(IjX<T-AeiqRW_H3KBWLGuaNmnOavmQ)!<kYSg>6(zODs3BU3lzSD7BOGH72i$N%T}S'
    '-*QA*z@7tGXA#*y*X?tRV0_|WsO$0Gju}QGM-e9k<r4vnz9?}jCCKXac9F+n(FP#3i(cDGc3SC~UgTUElV)mj0+(h8mRvSF#oS-2uA#8(+e5L1mG;wp'
    '{Bfy2_hkR!ci$1q#R_%}vctrLD)b4h>8~`Am+G}O+0O*SW&hpu0SR<;q}gx9=TWun(Ih4i>@-M4YV0{pKX<_=Xq0?FGP`22Mj8%rvzp@J0ZC>uF`2y('
    '(+_j8ESNe(xZWp>N;aN}l())lPk~jT<b8@+h%6(nB$R0qD?SF;=9ITsJlv6QKt$jaYSU1u^7XMh%pxm~)WBu6Z`qK$0g{wnz{5@1tD!rVCt+DG@bHj='
    'bUk0{mSQ4Q6MzmzV_&Q<qC^yK1BY3OOrau0phMQM2YBE${)ar1)xrw1BC>o4Mi4Kc=T$KPVBZ$`LL-(}ln9a~Y@ds8O8tuvu1rO|)IN0akI^l@m|z1;'
    'j0tf5qWm8HzZ!q%Si=fiN%NV36VB{?#~Jlo1ct2}k!%{cJ5ka!DAvKsf=0a0nB>N*bN6^y4kCo=Rg&!9=zDbBbuc0LQwaUG%OuA<n~fxEZILU<+}fgP'
    'DdDv)KYdC+Zj{WsYI@J$#gMg_b4cW)>nNt%4SvoM#M)e!O+KcWKF_aWM&u$r)n-3EfII@2HV1H~ZXKo?wwlgD#|F9|>2%u7F@9KV4_NOMXLDbOyXK1O'
    '?Mo9+Z&zxrP|01$9u`Yn2|M^0crLQoi=J36HAyHr>#bkYGZNMjZ6{ljfjiq7YSKp<O|1CGJ}(QQcFWwK@W$f5#ze~|`CB>7u^VuMaXhGyjbCdZG_VW3'
    '@2rj4f~l=w+QSLs$)_4C`r{?p7A&R!>y+(Lk0J1)$%S6vA7fKWgvR(mLF}Fn>-@lwtlSO};GdXe8$q%lj2#KW97kV7?oEcmDLIseXtewBY`54J@$cif'
    '(Q-4+=qX<XuAt>2FKOHoPd4M|^jKl1Kz3D5;3qM~+ZoNe#!wjX&mt{8ey8njgBi(L7Me!CUjFg-n1AMs=G4LA&y+N<v=)!7WKiv9o~{){N~@d#J4q?T'
    '6lV)n&nd>w-rl^GIGHovy-3!PYgxUBB}Uj7W3qTZ-N0=%rw|P((RCCI8UPO5pKCNaMMNENzpZbvp(!5Ux2wf&0qPQJZ%08ZxZ$*(B{<UUwc(DWB2F=4'
    'CA$aK&N?cM0-&m0aKZLVB>M&SQiU8@0ALh6o(NRy&A}NWdgW1Z60B8QK2@@5<nA!F98RE0SK)NFg;H5QEMRh@a4VH7u;qP~&1o8VB>Q~CInnygN2~82'
    'xukIbV0c-MU%`T%-OO*6H^uf=bq_Ks%A&w**LftBbk<JNYsZ_O^nS`j7hZ5%y_ou?)wi~wNWy*Hlo_cww=bmzxXc_1!={^lc|cAWIBKcY6<SbTC(Qei'
    'YFh!VNp-qRG)#btDJCq>s3Mi&IE+3HYDyiZCQ*Ank2~DkbO9~aU7Eux=)ujEmGT|B8SPcLo|g02$1{BzZjbYm%w#PfP0eK<k#r>nkkMM9m>p>n(m&wE'
    'e7iZo2kDjxMmGdE37-A-*Z=Y5!@m`8KmX(7m*W20hrd4mQFOn4`~B<ZDAGtqg%6w;7~6|!TE0Ave~<bm?HOujVe0IC_r!1kpKm#jm0V6;`1qWB-*4wT'
    '_&KU)!=?z2hW3`o*NsJYZ|Sx3FQ=bj>OW@uYKo-|aKKW#s1O?{WYVqnS~V`@1)a+5yu@B``@#drb#eEu<zB`<h$je+En(6wTTatGny8F$)x~zzzL6Ht'
    'ksEulT~gcUd#ql>jbfVr{$E5yD%hOo^unWr4Bj0RE+maojpm*@Zmx={eCRCz1#7pU){F}rC#K0uoZdOv_?U)WkLI<LX_Y6p<L!|$OXEf8TIuCuU?tH$'
    '`S8cDAOFt}#rF3Pe^E0NObUW06}ykWeEsmp-@pC-;Y&e&w!<I4{rch8kAEqW7OJsV-~^I8LOO<WaVFbZh-5)U!!sFA1H7d72@JHF8@MCjU+!iau6ORm'
    'xujHpW>R0KR7APBYFMXYlcG%xbU9lJeHZX!%3+qo>cJJ^O9`$c?|qM-C;oCbzuWBI$9loCxVy8ZpG4tUPeQ_NMnWoN(kldFH^N)kkn`taw|OdTWGDvt'
    '+9_{|XI8OZt<=zSB4Xkes+n{T*U)-~R2*=<%_zWC{yqNINnKP~S=!@o5BGTXMW*y_hjtsU_dRMAoQmeaOe#~&pVOO*zRiAfcsE}!cKDGn&0%qO2kHr4'
    'VbU=K<tx4U+JU^@dYq$PM7r_0kd0G$d+U?c5{#rxI$y_@G+YYOcV9(I0lhlcY;IknSMXZ<TwpZRmfvoQ`)&feR?hWcRZ;NVmUKK*G;hYS+HT;|xPnQ1'
    '2Zz8L_8H*JM@VjcOuH?ZS8g281m9s~^7zI97?)t0^DuXFtRWaj(N*3wDStjsp+|d9A{3SiOFB*;Rt(7^1&r3eiqVW=_*OAmI_~k~b$2@~Z@EKX0Qy^F'
    '6Iu+gS6mXmyO}CLNC3DH9<f<?5+jfpyeKTP4(yU>l|=0;&vq=_dv*{f2;6xC6A1UKHfdpFN>-?iY`NxdLR$hzjh@7u(Kf!^B;%5ElTWU+(6kB7zKW)K'
    '*e3y%`=SY~`HExbZUGw<m9ymGxvI*Sw#?HmU8U)$&NtkkiLGTh9Ul_kdNh_QJxLuk|3Pmex+8d%gEfD@*=8&sme^UkCFP}7MfPYdU4om5Sa;4RYCGlk'
    'iFvR&l-DBp;776n;APOS$|)fuhf3vBeRl}{B!Fgrj9~<2yr^s1FkJFPi;lX(X)PD|N%^3QvQ?w9oIxYXA^V@n24b5fv-zr8E4ORTK~_^T7U-b$mJ1pc'
    'Y*dgij_oBSO6`_2^s^R^UWBEf8imJ=twf>CT<$u6Xy@1J6fPS>XBDkC{QYi0ZWT!Yl;69w3nLL*=D$fKQt5tq@%V{mBJbPuKrY~?i!G{lQXIz%Bo+&U'
    'lsXd{=7yGL8Mq&on|E`m?s3=xL9|H6^{Uw5>Xv$vEtm1=+2R|K4~oi+U_y*lfTq?<ylPW_Dr70yMQ))9drgI0zSoy!h6@w^Hu}(NTPa}?_e~MC?*Wr2'
    ';Z>{KDQ4++cggQnJJlrM1_^qD-AdF573rCl%5s6UJTOSaplsL$O&XE7Ui4g`iQt_;Esvgrd_-W9=39{&KG~qHzl@2iCeV389EidCw!7}$N6A{lOPM=H'
    'FAAi70kwwmLCNI@ws;F_Vk-3o1W>!FRZz@oM+=X2RbC}{a}-N#p!T?!?+PT>R7o&7DhF3ytNaV>L(kZu_)ojX$Psl_kOQmX>Y~5i6tKo_r=6bd3ACc1'
    'O<-1!F+g$n_5g1H(e)7<;Yz_QGCFc`Bs*eZL|b*oF2?6>vU!vQ=`=mD6%R{S+U**q^<BJl(r???U)U+Tr{f}V{wzG7a?V?yV(c!wpqJ!^i&P87=!1qT'
    'aTD_v%|KPbfoUK<qJky*5S3Fd(|GS4Ye=dnpc`))o1d_34~&}7rk6ypw;i%>VzLk=*PG=d%i>S_vMCW7-=)snh<JzZ<_v%r(SoWh3zivzD5<XDrka<4'
    'xD3Q06`oL>1=s|!92q;Y&0`rdsHyjJX=Z`Mg!_08vhWa1Zq(Lr5c(XNBq=61KY-KhAn%!DZxrtS0CsC9sL$C+nV?NMfVOgQ)k-ow0u6Jy$YJad#tDBB'
    '!knBw_#zXeNK|Deb}h)EUABq9?gV>P<5h<BHGC)5w?y@gzY{^dqFW|3R4300`%usFK3>V`gt`K3!NomYO}?ocg>wn6J`WeWQY6B{46?4o3l5MzK{n?h'
    '$?#p^^d~56oGegHsZgzdWpo_|jo0&p#2`)Pn_xwXErZ$o{(zkl@TMwmv1uh5Z{m??FiuW13s&ho+MNG4cF7}S7FED*_5ja)$3z3jFks4?@ZYe1#%6hA'
    '2cW_pgtY1I7&OLCUxkf?V?^XGE}W8+Y(^u)_i)rvL)A5@Q&HPN%}8)0td4nr)oHywW{|+&UL*PbQ3)Hm$&{@KtNh1`G3enEFHT}EiTRUaF&839`6Lxi'
    'O670bNM$*SG$%z8{|9>B)sm<vwJ2Ir-4u3vY<tg!LGE!v4m&9!0}OR^wVZG#?p5^^(>j(|eWS!(XzZ5h_q)YCx%E{nZt}G3t)BU&Vh*R_lNP9zuRwW<'
    'V7aVMPp!se<I%#w`b;ncouaYq`=E+)##&F6SDNzogiYpAXSh+!H0@)Vo(YrrJ&zEJW60zcs>aeh6Qt&wT?U(tbi`J53cKnaIopBElIiUcvg;#9`g25P'
    'W~M>B{tom`94%8PqSsQv_Yd7<@=Tg>`c6(Y`j)d26JCT$G}VG87V<ct`e0k$?ujUNAB~E=vSSI|6pB?ztY*J2#z-Ez{MT|HG_Fd+4o8H#>%$7v+{f7c'
    '>TZF7boPIPDIy5z`s5Bgmm#<u!6gKDF+#`@(p**HVuU(Js1rhcF+!Uoq?xSF%p@Q55hFX+hSS=lX|2tV!R2$@q&cqLkHO^=-lPey!;it`v)`oIuj?O!'
    '%RfbvK1F@$r-;FVqwmnB@6ebZz{)>Dn?6G0egG@~0&V&NP51$<eEQoo{Z0J?SlQfrnXdq<m{scv6>z%Qg+UDSINvR1*wwjPq7qj8^64fL08cj?<fqo7'
    '5V`m~rRXTBmecNv+C4C~4D6QO-(nWW-(sJ;7*lIBSWre57SVnA75At+Ef<ks%r%YlFv1#_5wFn1-`<H=D5y-GWWtC{xV<w2PZ9{}dS#M-asHFY)?)ot'
    '7=$VtP-G;jrmP$Cp$^2V{1L)Mg1@y#zjZEq1#N8V8?w!MwaL*E({N|gc)f6jS<d-IdVKq3{<zT#uoadQBh+xQxW4Cwd<o?@m>8Cc9PycvfmUk8??B$j'
    '{bC8XAmwoR@(n1G#lwS%?LlCx%ZD(fIhA(M?UwWB0&Iab*;bFPm(bJ3Vt9gR+aFl})(+QyYl33NUhDZx^m4TF$xve3I<N=l_-$5){S^JjYT1$o`{?k+'
    'ZlXW9b=vBs;#t8gt2xoRYv5!g#qY}Alvrp)eVM*s)%K?NuVmc|u4oQEa3`*atDs5;Qs6e5uFz!j_y+}uO0}=5sy+q4xq}Iz$q>9~fwf@`9V0sEy~2dc'
    '?QXMOEaz(?m!htGSC}LFkH5d6<oSbOG`eUw=q2uXPP*X7s|!B=TBbl#?``zkx4|Z$O8Q4+;pK<J9L(1Jb}^gp(8!>`G#m}lMJ^~(I^Y6l(Ik{91-X0j'
    'Ws--#*~_kOgK!3$SeuWzxZ9W+$XR=AZd|RHSqO_y;Wyi*sFnz!U3Lt2DvV8aN%Mz|AiZjp)h@CLpkOB`UMgHbK_#Yy%}ysbr{>Y$TOD{~<L^QLKRoqv'
    'qGRKyc;>ZL+O=$mP8OI6shakP>fAOm(bxy$3`Z0ueru3YnV87H2MJ8Bk$v7ejm+^~1qTbo1mGROR-l6T-J-vSs^W^qHeJMP&Qqq7NeSLAV^!M20X|=M'
    'ku6Jq?}OC-)u0W8B*5<OSTiBz5>nUsFh%dbMiv1mXjWhh{|=d0VE_C>A6unN4ALQwf1UkSst8MyvDRh3&2g2`gPU3(^TAXxVvRp;bjY?H{qaBkUgYms'
    'rFXQ=m_Qf^)WX*1`7P*HBo4FrJG4||po(a2_79t%v4K}K4CM;I3KLMRki|_a-l892g)Q)QsQs)rD`d4torYZx1x5>5JXKcMwin=stLCT0ydW;bmmtnM'
    'Um9sGkHIZRF~mFW$Qnaau{2tL?$gix_|NMi`+XxdL57>1z`cZy-q~XFv>+Cy0;6GTs**7npvhQcWKL|%#i5*gN~soi+&hf<Ebxrl{0w)_{$a76i;(3h'
    'h@6Xg54F4H=Kdb8#*_u1GBO#+M?6fo&4P{6hafq0(H%?P*feVKBEhTEBoFi@>NQO&Q$iWhj{rz4L;z}={B(sQdAH{{)ea8-<$-8ri`_~krP#~x6_I6('
    'Xb!pt1849pad9ITll8iDY+LF^FutD%#nGUc26ZFZ4R9O7x<q@yR;DSgHeWt|ev|DKXZ5IOzLD_*RgcijAtDHAS+Ns_(y}p5MGbGFj+634QKC)~E|ZAA'
    'y<2wshk-<sBzz|Eo-*9>pK3yzaZG^?^GrWQBC9>T*nYmEfcI>Jo<{t)S8oq6b#Hn}mjI;BXwTj0d9EeWobmwT?}H|(S<mT6<j-fdl4=5#Cen#WWrX{b'
    'dO7kP*zMpuc9f0*#?gjW^z!{?$t=Yp+Tez^)=BeHtbs_KNZ>smy6yJ4m@IaW53}wW#tYg8=f=?;i*vSwNxD5)e_hvdNK~5EYB}ZP_${c6wXAB9kjoz='
    '=q^2$P({qDQ03EOUJ+_^Y<dk?9~u8n3Sq~rD3$gM&-SqtIXFQJi4t|0=780iOhOBLS*Vr4)peQ%`^p<38y-tx_PcC`SPq!;DRLe%l{Ca`kXXi-cPXbc'
    'T8LOZTv{$&@mwkbKNDIt(6Qf9i7E_cK`mj0T6*zF{arTfAI82p-JzStbHcERCnIGJC8?l^$!VJ6GR0zm>57YICX;DYZ~>&TwbK+;G6Uou&;S^m>SVgX'
    'Q7;N3GN<Kb=-!wr4y9>&MTb~8g@!&RVm!faE>@2m?C~lnDz-J1KN6ho4n(q@ByFpbRg$caB&lFL7pzoH>Qrqt@<g^&dSg{d?F)f7{E;NN#e0HQ1zi$s'
    'j|4Rs&=YP{IOu!U$5YD=pWY0~*G&bPt{KyJnwRFq-GrPLx>rn!xd(F0Th3TW89I9~*gH7Zw44r>*xLzQ!v0#mt>)d<hSgj?uwx;4PGOBh!1RCYeu#ZG'
    'xg`r6ed%H}8mn-|RoStzzFUB{vxs(XY4zvI58<3W`E#E?LG}|oAyg+tD*d_9pXc|>;OCA0+@_zKQ$N?OsmTHObAR&ZX)sm#c{urV9gK2ZG?gDizuSOy'
    'TsYffNkuuI9y2yYJGMs4paR)z)Mo_rMQrvW%^87O1=6}`$}<A33baYUoEAu*!liY=+43g@e5VrWTQ}jfgw<(sX_}fd(`5AnN{t2zw5J7JdWkem?U`wc'
    'sU#ra(_6Rpv_MD#X+N}QrYU$8NQ<dGvzX*^27&Ay&a91E4*{P(b$e!(Y8{d`V2Zl)34ummIxz;Y>U=_=)zv48ma$w!z=tBcwYIy>-F&~GY9%3Nu}ChD'
    '4$U`e^%$4#KE`_u!~;k<0P*sH4d{s}brY*T(Mi0UgJO-{ah3(>ya!?mU#h50qV+k^>0-H9&)`PA!tBlUJ$1F1KQ5>-k+g0=K$ifUlYq;`WWIZh6)vQW'
    '!yHeE4`aTauesZRK&Et7PHvJ;?3^WwzRkN1FTORVisn5{C22L!R`n*xfedCDw*V0BG0P!V$9}IoD9<bgAghzdj0e5U`N`%mg(2*_brD@!WKO-6sSin;'
    '<e)>JYjvNZ68PIJh!IA4S<`PddB_G|ju_I!K$KciWu|m0PzG71bjCMY^uv>Rjg;<&(0U7Igl;0sM6~Zvxxx023#yVQ`ehrwMWHIFR>FD1pt?-v6quer'
    'kq%87$z-vaVZAXM?aCnW)Rc<TCLDhuj?$$lOszL{V5Y_RJ;~BfJuARMvO2|F1DSsgA+|Us$^`>jPgcun!4|XhAH^JDyMEU_XG|HLiVE9y*#vi9R9@}w'
    'uwFlBdQmHW;l_*RbrK_n##j>5ot4p|=9N3>?Lqpeu_+Vm`Da^RTe(yy?A1zEi3wFqY#3z<XkXeRo`6<xs?<4${h=f8KwnI=xx7sZbj*>W0&PAjOrYK='
    'ra<rXMk!4|NLy_=)|nqLFxE+G(w?$#1}#avP5q#)Ir%8{qv^v`%LRS6ik;Cn$*&<i6QJb31Wu5)BB`skV$zHF$bQ9>x%J#p{KbD0!A9Hm(mo_N;0&$s'
    '*@8>F_nRdJ#ib>gX;aEP%4<L_NLLfCfL`p-E3bw2$bKBI>kt3=KcvSU+wsw)@n@z<>X&q$f)?5}w&B*P9l5q<SwF9wn#J6d<)=?YTVSz!x5ZN{y+RN}'
    '5U+oG0*w^zxU0=gjZG9Cf+<R+z3?1Y5NW?Y<&U^B6-GqgWI|$eU>~RBbu1muXeBExp{rt*w72R?scjkENn`HC^<w@mHS@p_4Yr^$=}^G71nF+6N6EXh'
    'eup9?d0BPIId{x)lA0_jy?LBCn1B~yqXBVMJ6zR_iOaf=(wtn!oMA;Q+|Z$@2h(wrGfJbNPD7L?El`v#-Exkssa*OJbz3EX<~UG!Kj+XMXc4Q(Js*pB'
    'Y_Z=?W(rd$n2~rVm26-+7apsKw}D%z>w``2BUq!^0$@9}Sw}~_fw9Hle9@0Q_U?0f$4a4eu4}RvCep857JHu*4J4{~q6Ei1)Q$@zR@7og*}PKu*ZM#H'
    'y)CJxROY31`YO#ayT=1;Rk)JxSYzTGoF`vQZl)^kl6ZYiJX2Bn2<iK_WW+AJd(3&pDh3K%iMMYM6ZK9bfFM~{=hMsAhJsX_TrTpWxdz8&73bvLn7ke='
    'ZDU@M4`cG_%^UZMe0?l`bZ1v1cXp)s9KEyS8i4y#J|em*Gpk^;m8RF5Sb#*e(y}+5Rjns5d+y`m@xXqr4Y^XIqWtaV>g$K!fB#sR;zLpKO;cb;08HPX'
    '?<uXSlL(lz%HfBK>!sPUd=^~`HNi*_%h@`;XI+oz-gmp0NUb`zbE<RWV$VNPbaC%Mcs}hWSb7fI%N1AdmrwB|bzXbLR(_0TdDSi7Q5JX1Csumx*K&#x'
    '+9YW6I&`e$83G^ShC~KQw*UR|Fr_T##paQijefe9pH@<|<|?*Fvcj6j#eR=<iF)t}e{HvYM5~gDY7|EnvFF}E$2c)rL+XT!;*Sr1e*F^j(|qd5iwNyG'
    'gsf2b7^G4V6GlO}MaPJ!oT9$2%W6P5MfVGg945h9328d*g#u4&Z0|y1&+R_uo`IF}+)JEU#iJ902?D_h>*wY_(JmZ+3BgPIqYjz70Y}y|+}$ka_bGCk'
    '1Vjl5r9F?bWMaCv9Mhbysjd4Fo}}eB9A7+x7yA<*L1?JfbAk04h#5B1SyQ98BOJBzur~3P37Rf;(~;RB;+2fQtvuqMTpv~w44t8tm_i03DeK3VPFHIi'
    'Sz8809!zy4ryqt+Re&@gx~ij_Cd(@jugtu6WfxXpB<I$}mJ;W~Y|_fC7hA)Wpa`cqCbs2G7kADEy*`<0#@ZsQy`ZdGM>!{O@Fbjc8McNSiw!SBJx$|v'
    'n6APtUX9o_N#Nv2f?<)`eUwv)-|Ly<R%won9Lgxxo)yDH+a#v#z5=P1XT*Fi7Lu4dFV<vYm5OoJF?T7|91LRiqYAeFs%$W}=~ZNdRb+$q1tqU0`I3;4'
    '=CRmJA2&<|C5f|^eM@2$M4O6<ZHzghbj#8@B%fomn{U|M!nX<83i2tm%CxYrWQ<NXhxOwQW`#PkEniC=#Rn*p2Up?_Ge2{5?$Q_^Ax%0@Js?Zn@|730'
    '-dCZ#JeSvKE_)e2w|*l{G(<@IBar?Gl7<HMvHt-$do;q@sk4_`X`MMxb4TysRP}WV({R!F864ry9%3{>Q!_a1e39Cj7sAo*8PhxgyGqy>tF(BAVuz;='
    'W^+zWnj<9^`jq$jr1*?BWp8j_$UGPKC%g`JTjX=gzt0sPtI@gfCRw|JDV@bJ8SSXkruxaG{a1S7Hl)J)I_KbA7qi21I>+QeB+%eSs(&Sda1?Els68WU'
    '!;xspuHnwqYRE)pax!07D+BXmJCDXC7Huw>ZW~Q6Bo$b5jy)!;rfzp`q;JqsM#^}tVPDG*rMOtoXV#1*w4ozoB*J7c8o?M@o~1qMq@%nwb<l+#Bdn{r'
    'sp$!uQ5j(~HrQGJo(!4m)}0)~6gAK^1}%yOi>#@}HHama*=&}FUP$4m#rhuBfcWc!lN4q!(@dq4*qb*jH;8iwA9aq5J#gF6JVu{&@WTHH#t$7HCz58@'
    '^Yy%-|LTf=`lo+Nj>_PT*KOjP3vxE657>2rnJXejP>-W)64ULtHvo4wabF-y0sLXVIr0vKplUPRUv_MT-6{oL!)eyo4hA!6DdrdO%VW2ZGHD#4G8Lvk'
    'SWzV#)FzH%e))+@?UnQFNh*`V8rscl&K-7(Dwu28!=aK`vd7%fi*Ffs^3DpiT!S9nZ9jn63=tb$_5O6gLJMwX-8a(z*n6f?yKEf(WU(4Th6LIC0Ay2)'
    'VCr;*viyEbYsWC#glT^Ovl_vyOv1cO9m%tc`}0tZz{jtJJE{JkgGZ{hBlwU;-~3O)k1Pz~%?6dNapIZF?yA@)c3n`KTMUl{CI0}Gii<O2MPbkJuQvqu'
    '1vK0+*r-D%w4yOr4fpSO6~@7BE|pR!uS-yz;;+bitg<^AM&nZQ{h7DXYZ`m^s%)Bxc2qX${QcQ*Po_B@Nl0kTrD&9r@?=NeOFwQ?noH3fpfykS_+x1i'
    'ZQurF7olP29`|IA-?VjUY~fNg>}l4#jC?^^#xOQlo=4NKc=|h3Wg;tTNI|?1@5J0y+O-w&+~s&DX0OsxQpEF@<DHnlN?V{JUTSs_(Ije0u{RqssYcrr'
    'kWXezDphAxrO||4{$OLjccna~0IRqN=t~oESu%lZvwGraR>&hUN~=P$KkHV=`bUNExhkW_3R{UsGMW?~NmEOeOgqk%XC+-ZkZhBrIWOt^l9E^tnrick'
    '0w*Z|xKv?g^d6|nM(4Jq#|u#+MLN6?a;1wQ${AHGA_?&sZ8Q=r4;0W!bJfwC{z`ZZ?z;7YypWHiNFIa9WPl1<x=~j*IfpWTH~C(UoZOYxBZbOxBb;-j'
    'R6BC?kYYE(Vo0RoJsO-Vy|SCv+;xg*5tsLy8S4J=%9X;?Sf06u>*nuYz<hP(r>}zKwarM3!V}W-0eJ(WzbcX>e)?2YWdGXoQ<BzZ1<d(yEG_=LPoJXY'
    '?X#zg$pHET9RZO?J938NJH+fVtY!c3<?D@+dQuP;zFQDBLmKZv2|g(wM^L5}s5U{7n^Ecxai=_Sn+*9x_wZg5p?dX|*FBC@d6ja#?!dyI2~(x{^y$UM'
    'ZTd|FtrTa1n$A6e63KkG-FMRidgYO}KDr8c`a;tLfd~>uMF@KeLL*QKAtv6!X)7O~;!OV6{TCgXTGV!DySS&eA#lPDy%<%H&<<=RC!LbmNtz7ZWT$2='
    'Am-DD!+TK>tDuTZWxcw6O@r>;jIN>7pSS7ft<<l@o)g@qWKT&kNRV2vhct>-VuoPjw{S9ph7YhPlNy}()7f22t27l6N#%e)R&`=`&H4ewiv+7i*|x>@'
    '9864-N-Zl*FhPa8JnFD<n|IiFdzjy2gk?$f&Y0B@V{&%~%)6DQO9?y~oc2sUmgFJn%Qfk%bF%?C>-H3(>ZsQYydPQ=iT(^Thu_ZkZiXFQ9tEn={-RCQ'
    ')5Or*wV3D1G<TZd&zs%Zb!jz98fZAX`Mb^X9o32bJlQ;_gI8X>PChPw+Q&sTNb=IPX5d$~F+SsY({wd>B93s$#j3^I<hQIf8y|#6_!Sc2VQ;z&b$0*d'
    '^Y5R(eEh5U>)V&Fzx}Gobbx6SrX3p)dkpCo-7XLN2fAXq(n^(0?ZkBe5FxTx+;`%xh06%HGZ)R4XcX=?u%DQq0}g{(vkL;G1!QpGxslW^iWTV>Q{sGh'
    'E~TWxB{HGQQZnGuON>mWlz@RW7t0{w|C)2k;Tvv|gnYFckcyz6^hOY*o*j3z&AMdf=T=tSDNd#6V|o&0s#f-ns%FCIfdih>i?YiKyI#PpWaG(PCl6o4'
    'DUm0dAZ6bacGq0|q_vj>fvV<4^yfnLF~>J=jh$spC}ORM+yXIH6iQj5#NHN@#rsZ`4<vDzb)YiG8?)x<9c^r`rkA|4JLIyMNs{6EKmHvfmPPpV^)HpI'
    'l4NyO(hnu;BpF5|V;0ZoLz^U8pBEiI)$3xe=p>RElN`kJGiGJRsx<SQ^9EOMH`Aov>4y7q;OD}qbk=p!*R{BPs-}D`nE{T%$LYgSpsKFL#I73iS~kOk'
    'b6v+kbRgU_h;_Yq5Nkk3YbF8UXK#%Z9oCwbiSp+BAl92lqD0^0p-a;;7)&GOZUEMf(rUs@y8h`|s4n${X{2hqYHVr71m9yc4Q||Bloi2R1v!dv#tMiD'
    '5&igMZUEGcXhl`aU7Tbxr;9bFOlqX~aj`}AqxHR3NNWHXK8O&qv|+@lel)R~fiop%VVaaySv=cb)09&{$vJ-vCA1G!6XLP;SUe(C5ni9EYvxt7B#Va9'
    '2=;G|j2$>oK`)&y=c}CUqEXiYJb+eIZgOynINgQIxd?a_sO>B3Zr|O(=Q>($P3c3z2))L)`q87EB=l0^S`maw7c4WDoHZg>pt?`7>UOYhuqsqx+WNIr'
    'so8TeIouz=j*;`IuA+LlA(DtCY&ksER3k%E>xwn9A}S}1(*vpXF?)h-t#rNaW~5cZ1~X{}(%s$)W${36I-l%Fs>-C)G2Akf!LbxrOtT%Dt+p(Cl8s7I'
    'vfq*91MICWoHQ#G4IvX12PcSfZOl;gae+NVZ0YMHB^0_Pr^JdjpYica?~Kw*{|?`9dvJ~fa{smSvCZ`yGrh4jI+=yT5GUYe%+TRNJTb(M!K5qz^eT5D'
    'm{cnrgVm{lR(a{jNFz9gQew0a$z_h9l2CF${-pBC6W^Yi^h|!gAe-Zg@-<{1)WTJcV`wk^a3><ufM&D+gJwEz>^W8@VPYHH2XpF|nrVz<6`M`t^)P2&'
    '$Sa5W7;scYxkqJwzBf&k<#(9VYVB&iTW0+vL9i~h-p#t)Fc9v=$W{=4-x5<m?OUaKvV1pfsh8<yzPy#bHmEXESqyxE`<Q9Hj@GwF>*k3|C_r(o+aY|_'
    'BzQVsB52B?Ne|r3*PG*~+WsS?{7@bdN@;OXVk}%8Oz8nK0lZ9alS{dLLZ3Neb5UA+1Uw@?A~rY1PV$Bv_e?<%N#r{Y)fMv<oIu^KD0~C9Q`A`_!b?o2'
    'w0LgW?}K2TwpCohtu~#&YDA;lc6ymGJ8^2q4HB(*K@A^U=c9Hh7dmzapo7WR{&h(VSWpz9hEp-AP_2=04APgimj<6|qFuXk;R&0>#{;3YdRT&z{YpLt'
    'I#c?0?%<|xt?2A+?6e7s8%Ocb{oE1xOmExtcpFh~g0D404v>iERB@<H-E%f=?iNJSBOCUSt9|BVh`>8u`@GJtRmughh{!?(V`{P1iRR4p`Iahwz&yVC'
    '^W(3-{rb81_}8BwfBEgh?+L~x7?)tk)poyG&Mrr)q$sA2R47Np{uvn~;lEqV-)@q4lllU@Ccf(Mgha2~lP=<@*Kkkt{Nmqsi#>YF#5X5bzA%fjM14-e'
    'i3BSbEK!|mG_4O{QrSqc;x4@&O%t!jR?8@1O6jJI{-?7pEiiM7lP;|`NRB_={q?JxBU(_gf^4`qLYsE2IZeeznYElzn|6Gamq)tM^2iP+Q^RVlGH`+F'
    '@3NK*ks60iI8d}V&(VQ7Fge>w;z#J-3dv#2J*V?nQS?P=xF?NTt3#U*QU`ZAtoR+NM;qQ%&q5lsc>?9#yud%kq;pEk;@gth4S5El{Xzkct+g@oI=R`Z'
    'Se?YGp_nLrQ8JfP-aO<Y9m<L3Ir>CB<SV9zp-RxhTTsPt)pYO3%DDR1-~RaJUq1Y;`12bcRDb*Bw?Dpq`|_Xv?;rBM+EWH>uqu-1isvPkUq;^~`tZ23'
    'qRL9785{0eie6w;2~%zGB=KY!!SZ~!EQL=9w5~>FB3luCZ)GZ%@IVR}7{`e@|8B@wqsv~<oJnZy_o2BFwDj#GHPXds^$1#865QMKa<ujsE$F<pW=M(='
    'GiAwfAODP=-}#Lk!;Sc<VJvMJF0&}2>bES@T;3=7_Jlk?enhqrSy#Ffvf0^^7JI^IwegVUj2OzlqWfx)d6kt%UQ!(i298|Wj<ZjW%YTJ2Wlw9cMs}}M'
    '@!PiI2CbdVm*6?3mgG<4B--@>&E`qBL*H6vv}9*oD~1a`rup8-OqPfFa}PNIOT{!yK-0?9#<j564CeQHXE*4mxuF+<^T+N3e-Cb8s(2?2<=`UEs>Wvv'
    'To$W2+}@o;F2sFyw?(mHn&TKRr8i~GOtc492Gom%G1b|8mLiBzP>e4a0GS%ndO_QDV_4-d-;u8Nxa+3KYukfpQo}P(r=h+>H=#~^#A1Py#LciW(bj5b'
    'KNHzEKxsfVF-4nUVZfHaiuu%IG|<eIl~NI#tiXfGJ`+Inz=Pqq(@QK;N5TR1&=EzoqG}COq^5<5oDqqES~<MwRYDI2cE-G7i~qEEe0(0*+(8U8X1-+0'
    'il|FTd!YAU#E2lMQUI%^?+HO;AT_6v&UTyao@);F!kE2}fo=@q{v?7D-E1p%=Z*`K;-ey|mjk7dkBi8~-+x?ClGl<mm?(1r6<jn(nUav%=4l>xb4;!P'
    '-F&jtDVGESm~yfVH6tP>L%^X&TpmfwOn!Q#dfhzI|CHa!f;6}X0o}&@sDM@YsrxLpjmJ3bcC*FO#{rEDh^v`VF?s1pzsY{TDB!r5V~$prF9TcZNyF94'
    '=T+$C<{oLgACXHR-it!>b8JdA6FiH|xL9&qYw{lD&Q1>#GIhuo{lL-n;$ej}-zJkarh&sxc6V$kx`p!riJB*e-2^qYO@YNi_~oxqWhIue+FiR0r@>-q'
    'BL_{*`8bo`fBWm}$1fO4>%M*c?d#{iiv6FTzkDrz{aC#H_V;fezkV$K`uy!*KmPIC7s2v+&()`#f=3QHTJn8o93V+1NgcDZdiKxYz)iLPAXuq&?0wc!'
    'fayKxS-VyGpkleUC1=U>)CT(Xv8E8B;ugh@G!;n}8kJ0z2oqwB1;<B|dX#1Vsc^%k&8EDoHRa;r5KWuRD<Q1>u$%8UJ8DCHWy6yxzr7Kd#9ze6W<QYv'
    'L)`n!_N0YWx>FV>;K{AumP|V^g$oLgjAu?f!1u)N4!sZP^`>;pz2B#8N<KRxX}b!43rN&Tksd1$>S^&IWNB@YKaSr<s8KitA|FAeT(+K@rLh-nKcA)W'
    'R@7{`cr~&vaew8fWCgNONI8zNoQbVMSH8eMVClsC7eJvz3D1SZ@p_idl-xL2xv}pTyCv-8Z=e4B{Uc~FAHRGk-lP+XElU8ZP64ej?Pt1}EWm90@Z~qy'
    ')$#N7P-;f8Cx??UA79;l19R^~@yF+X{p~}B(v(`w#AK67_uXUg8`|8c{r%KE$^<u7ZO5ujPB_^w*<Oh)?ocGTK^rD+j@c$TA7R!q1;&DZskTWXx<D|T'
    'N?(&nIjAF?9d_M&xY=ZDWxEYohZO@aLiMwjq2G8KZGrCbao+DcT^P%qr{5HC(OolwpjCBZjw4HJ1TD739tW!un6=uy081{&8*y0f=al9bjgx-A0}xWu'
    'W?xZk7q4;5!YMjj?bINdrk6@p>43(z3Z;m*N}4C%^xQQH-we?A?CnW3$bFX-R!M4_R8`x{QBfCSVli3_`w?L7ugFZ`CL@51cLY*T#cJSwei$5Xf?<r@'
    '({P0|Gb4=TC+sFM*>#IG80*E=*DoJ_|M2_AFP{tk6Yi!+;s6lZK@(5vjKE4tSNGYuWQ?>5M{TD(t?ewStJ&h7J8-e~d$oEluo=wk`Sz#7ZnIjvkv(mU'
    'Mi2v&wj}0k+ytArq0p3_N0qNfpN+<?&4$T-tiZmD$>!%}My}4=A<;?D`i?D{MVwqSf!ziDW^`&N!4TM?9E-@azI0`gL(xcFX?_UG=1{aJ0y`qUAIkTj'
    '0z-u#f~s<)mXYcoLTd6sN&Z0-EIV-11<7W2`oJ!^O--)IO}c1v-VKw>O$}cdNd;jaVv;%XY717Y&n->D(s6Q5wK!4)5QA2Iz*YdW7on^$$u3*X;hNko'
    'p5_!`JH9FB??Lg7AYKaz;up_ipxVtUodjNc(!kGc{&Q>u1t*De=MLoPXbv^0*eP$2HK@TIN~w^u&Q*~vEEw(I)PQQhPhD*{RJN#ga||Fo3H3He={*C7'
    'h#MXv$zn>4;bn<1JrXF#?TGz%Tx+~qtmjnx0iWpGP3Do=b_Ye&MEyGWi%@r)$&x?&NlaZ(b@fVS-*)rUe1$|s{@YES%M%Q<#D|uXZ5ObqirL~GTT~Q}'
    'bL_W7k*iosf}@#~ZR{fs!n6b{<C<4*tX?zA7cX7cubS%Gyiqu<h+t>#dB_CEcODpFSD_KGn)JVlJ50Zp!@!A;;y7`l`v8bz5dS)eE`w1V&p7S^99@w2'
    '%WgFaCf69~Jlq*`{l1e56(qz|Zc}oV%Z6wq^{uUfDMuZ-Sb2^dU|_D~8$yE!ZPVIvBw(gFuvi$VP2falGTKEGo=hKhwLqG(12=@~CK!a;cN1a%rsHF|'
    'qE>FHQ4qj{<V~vb64^Pf&U<|KvA{c(XD+ic5kyK;Q8lAT<3RQZ$SF`=eg}|00!mX<yYoPtn~?qjUl6H)uQ1)?O%GMhowha^*_H4wQF~PS;Ja3H*HOy9'
    '80jNX)xI4U+<EL_r;$#SQ?;$m7=`*m&F!Qrxm8|IH0ws2s?+aaU5Vs_DR19dSLRZ!e#W}-=PMT5jF1i{3=X`ayHOmQYE4WZNBwhYUs4O08swumcZQ{y'
    'FjJ488AcjJ^aDMD;YJ5O5ia;_RBhU?%%*g7RH#jg3DjE*W?*?5cb%%g-CSY$18n9)@kSwSf^-)`!hh_}VTv9CAEB&93z|5WDqhjP1ng~e3&N5vtf>V`'
    'T~k;}qp40bk==nBb>aLyF5V$Q9Na~lU2KgB2cdoh+i)KYHY*Y}+@9laS+^yW(&Q?xMIVt8e(y*nXG<!gR0J?q>QYiMs3~tl77luy@4ELHC+!9`n^o1i'
    '+vS6U@m@{GeNTdk#rtyZRwfoa?9;tR+QTM>K=5h?fos#lsI6D7%bEkgeGJlI`3)Jt%!dHt7Y$8awI>G=b>}z=-veo48!;XQ?XZzipfwuDFzt!ASj(Km'
    '{h5gn1b=OF2%oV;hM2zj(4hf}hTr^Ms#6U}rjD-x0&v)GrnE2H9E?^?`q(Wqw84_cLc0zO#&waoJ}VOed9fv&ra=xPt(FcklX|zN5*>K6B{M3f6;48J'
    'X@!NTYPn&>sL;^wVM~W&Z=I3a0n>1WcB)LmyZ3;_hOxnAH{C6$KL>wYbJLUsQ=Jz}7kdr`%zX}JjcpRz1gxtSZ}$bKyTr!jx7@dUC|xP(9nT6$zw7?_'
    'fYRu6gAKJ|I`?&KeU3FEX!`b$uqQyIn?euI$x0JC45Z4W?qL|GXRm44c%ad60V9qldhMQ>Z<}CTYx*GGnl#hl^gyEWBtx-16bt{#D++C$ag<olDJV`*'
    'YV|;JtMIc;?{p}lLBPW$uc7vJUamSGM#mu0df%-m(k5?WRNJb3UA>CI4o_7D_w(huPLb69tW0?_#!qg3darJ^*yGmNEiePS-wbAu`W&nk*B*?MtaY9s'
    'wQ0%Fmmu#;D5gueudi)!u<qI#`TAPX+15<4B_0ZPS?tvrJxJGuFZVQaRwm6$y;EB=GZ1lbWpU?cEQDU5(RuaDhc6$$e)!v;AAk9mk6%8>Euv;~+i;7`'
    'KJ6>cG3qrbfeif%$i*Rwjn(>cM5u;oV`B+pXoz*shxI0dy|N{k0&w@qz$)3GZ9FLO)dCijcaMjq#Nu&6_K$Gt*;~;1Y<c~Ft*8Y;(`{ZaNy{vJH!y##'
    'W`DL)mBwBf`qZ1Z(!dk*O{0Lsz%V=hxYh|BeQZ%t&x;Lx>&!1gXk}d&@**|fD0S&#{Riog^tpMLOXcn|o}gUBN`V{!kR%TZ%e|8Jf}bPbqKY9&^4^*H'
    '0t{}elG6N#ca1i5Xd6D)U$J>bYZYTxs9{3t#sEuDnVGsQ5F|#b)l8ZeZf+?Z!$=Em;ym>C3_EcGgRZfR{5^)2!C1zw1u&vb3IVxNBY;iY9y<Wu9`0kq'
    's=nVmW(9P2a7I)ZKOLSku2iEDfWsMJKD9E?lYmtMY|a2KD4|RmYF&a{X9uo3*_H$wk;L94?+-DH1}QPss^VVIv4M5Q)||44DRJ~YsL8UjL0+vg?x69y'
    '{1mL~B<}R_{#9`~hE>?7s^3?~&PTb5hV<I@HBrr2SLsya?3&l4^R7%NTc;=cOP;KPS&egf(M%6sxlIp#RX#htoVNLf9Y&+)W3OAT7CSA)iOSX>aUu_U'
    '1CCdmthzYiFjubhZi!d$F*BqB3fXxH8fCxUl9(AeDHAm*&K0K?^I34mY5hO!@TxrSk40>ni$_WLBExBAU8C}iY@X6uPI%X%4Q@IgsmaUK{0XIOlB&j}'
    'gpNl@sz9cdjHkBBk2zxV%@wsfZoz{kC)5@N?y;u5ju#ndJ>pwHf~<ivh68JfaNu$NgzP|A@j)yUf}x5*%l7|0#b(eCx;gMcc=FQ;gb+xyE$@2apE{-p'
    '<z#@X{<6x>e=ihJZOkq1Unp623wH7ea+{zAzJGfyX@4%oAwUOjGsnQY`5GHsQ#VBJWzVg`#ck?>>I8wa=Z0=q&e0e&j#t-&s}iv>TY9fjsPHaH6e<C~'
    'YMTnYoPO42s;=pMPFgVgUUyjl@zi-Bk>O#EuGWe%9n2ju*TkwZV^=Hm@RX~^eU4%K815`a6EU!V@=Y<lncO@f)%%oABDDSlxs@Z=FpsrV$E>>h&F#Y+'
    'H(IK{#g7yVYmkPwt?|w3T6SyYTsySU_KVeGiJU{bM97ONbsY*0kQ%+UD(<`5xAx|{>0*O>M)m>Ytp%^C`gq2gG&1joc+uX=U>`Rt{K3!g7qbWo7}t>)'
    'hH(v!vu+hPnAdTa-TO^1MVFA88mC+t$jT}o;cidd2l?#90g3BDFsxOwAi&6Dq3>{Y+=ERS6Tle2+8zL=Crgq90F1Cx%%5<ZQbEvQ!b3hWO<R^u`*_;2'
    '$%nF>ud&^T<k?X?jZsjjkAie+X?iIfu=(f1f}4_}QyTQHT}<~O1S4_ggwX0ZxJr@i#|3Qrc{;0lFGE|s50aW~1Bkewaa$C(aQ^Lfa1n0kqPN-Uhe)wK'
    'afamUgya<|I`CqG#TC1CitFepf<J(%_NGu&e;w)H5^zXV*e>|~d%FI6t^_9>fz(w3$&&84A}eK=D?oibkfu#<X!>j~De}XXZSUiQ-~g#@v(Z!uX|Gf}'
    '9Nk}T-pvK19e~)B>|(j-QZB-!gxRS4Gt2slQ0k6XsHEOi?BkOgjLbI4Yo%Po2w|%MgiB?TCbl|yl#l!zy?MAp9_Mclb8-P3kIM}}U6WqXxID9fr)QXG'
    'TOkT|sqCY)Ra)cDAED%Vo+#fAM&UWdH_F$;SGz-xNv?cw%78dH0MXZjPZ!E;>uT#jQ9RgJ-uCxfD2EuPP806)a0Dfyp{sl~+3b56h>SG+`oU;$EII=R'
    'Uq5}dye$J`4MtTtwQeluc=gB<IUS8Gq8giYzEJr+^UES0h!3}&DV<3#N|qO&+;@ld^CB*Bn(Uqj22XA`MDlg|kX57F=iN?i<#-5aE(F5(dH@r$lTNRa'
    'u6K<+7m_(WyiJB+jqDpEHuQOQ;5fPbNGkwX6RY-Vj1BYLZJ96#i&0sndyAh0IR{r;CMe^pw;V$GH2+QRxH?Q9HnAjmcxxg}ULS?2U?NH*Wl)cY<uWZQ'
    'Is7<{Ii7pI?CT}~W_{@utaZb?Ml#%AxoqqRa^;$|MA~+^exufNV(m6ZYqx4t!AJF421=06g($W=o&6F7h9KyW)YF1P1;-$)*5@GlVtPy%ZdMqLKW@`P'
    'sY#Sn7c-^qrH918kLs;@Aq|tM9`Rz$0Ass>&)BbT4&G|$9}X%znwE0ZOj<7n!D!W5r?3Dn+sR@XbC+0Uf)xm+<imV9-Mm!9gc^mN>VBrB!{o7JPW-$%'
    'K(B;|=pZJh^;NRAyq;~?po+ZORGP$r3<+EBF&s?1lhe5*zOFlJ>w;+*FD^*7i+6gv!7*Yl${9-pZsR#)-syrf+~;>l6Eg~~joq^FznFGj@~;8WR~IAR'
    '?KbbR@6RZtrjcOCbBw{7swZs8#nUAbb5vecPT$l?zCI~0yxGhS9#nzynJr<Olj6Ji-DdY*X|xj(b}$yCHu>`ESABUHV-?A^Ob#{ox+gF4QJ|pG-Ehg{'
    '(cn3hdwTNHZO>hPP61TKiRnwq3%BSw0dR8qTc70pY55E0Kf8%-V@G;Y==o9kwy8bdWi)+oni-~Y9qVe3jiKN$*}#3d!I*nIjIJ;X+XY$0m3qz%cdb1>'
    ')*38m-myk9Sy5&(%hpEKB&G|N&=?Zt6#r{Nur>1sOjc!88I^95G+loi*m6?&m~*0V8V>14R9V+FMlOchHc~unxDF}al0!b8`MvZ9Q%Sue>x#WrYxE8k'
    '9Y#Mlu{|j%EpV?t5Gg5=r@w!W$9k${K`I(S&B9z&_z_g<q}oGN?s7JBzZWq^f)y@(v)cRs9Gxot0oDo8{s5w+<m8ZPY1HmVAeC}E1J$Zq+qAEz?cN#6'
    '<+_FYVR}>u<D2AhXnk|!NrM}Bad(%6Lg5iR?!fx|vw_PW8)q0KqLh8eUSI@tTv{waeadiwaUnC?&<AYofB!G+jjbAYI|Ic^95k(!8FZ{$>eft~(#WPX'
    '$+l}<NXLOYK3TA1k7zd}i@yA%U4%$OiO*%CF)a-xM*Hr|S2zsf2@qm4QTj1y5kD<*&oy!puE=fvOnLigDYR9xC<R3qi({%xHCjgN(B60J?qNfu#t%gP'
    'Xv3h?GZ-(Ii>pZA6~xzS+@_2DI@PAWG(%vk$$yTU!0<<&Y#u3b1KXd06>|?5&eN9_IXx;}OSy~sHQ8wN*pmyxLR+l25Fsh@L}x(eETSMior<Xe!&+*q'
    'R)^hvbYTiiDsp0uVPdw24+A#}noM+HO}Loiuq{HHcQg%w_^Pxinzj1UV0LY{M6f;Iy~Ebtt`cv4u<C2|-1$XwTe5l2n0#K0M&g4_844vCSZ8EY%Yj_J'
    'O7hit`Pjxi7jBYp^UA}epQQxQdqC=~@!Tn;AJj?q3Kx$hS84ii&dOEfIl(9M%Oo2qlrMo)9Nogsr^)9W$h*Z+8WeVhoG`Z#r$XpC8=QAZ-oBnxaRlWO'
    'l!n&5t(y~i4%W-uzU?Mb=fw3nqORCdjQEhOV(g2EY-|S{5iuRDv93Zl!}R}W?ah|sNU|)!Uojbx76D}vgz0YjXmB_HC*9$m?r;V`nHN?BJBppbU<w(`'
    'M%Q1A^k2-=jPzSQR{dDtMfIxM;vgbrQZgf%OxI(ZE_XfWc)42cnV7bfP1O|<63a2%TrxFOjD`)5Nw1UY<w)IxDzZZCE{@usW56+i%Q~~xtI;w6U@URg'
    'zTOqAM5&T4JNK3GW32Wov`EPh%;j=UdKB;gD|I}+TrI^eRQrGZ9|JzQY`MGQcBIy2d5Mq5&E|3LB_f~O=im5U=AIm9W4Za7ZD-K7oL0(|F-*-=c*Eh7'
    'cn-0L)JSa&m<dNjW)1W+u6?SaDI9|!!*a86?z~MD@CYd1+mVR~6YsLG*^rYs!3sN&Qo}4*zK?I()^!463W)nLUt(SjVsXGND;>&1$|UxktM~(Y>_M)h'
    'z;xEe{Ju;6!TQMEIGal}=a!h0Wa$EU+i+OzxBLOpdtPq@Aru6P!7LWb!xnR2*ZrQ1G8t73>dGKH7&OT8_U|wXz3IvPpSyV3%se0R1&2;g6x7NM+4Bg('
    '=Z}<Yr?$h;J_V^caNH|HW0)g2r(9D^nOk}9UlX5MqBE4Cx(w}xox~}(k*abJ^0iEkcco$;a;)3XRl$Cdmq>|1I>)E=A*=O5QY|isSg~X7ro^PN!WDsU'
    'u&2B+xQP!dR#Buth)69ORaK0*Q0kSq@@Z(a$$^Jw<3*_(x=5Q4*p(9tt`l7~#|-Z~+<LfMnyD;@;bP2|=Xu=~nFdx`%AY84!fcL*#c_>lx7*ho$Z77u'
    '_zve?MfeFYa_GhAyJvJPxJ+C|$-6M|2P`~kBz>L}>^muTPI$J}HGA+S^7jdi#)q$e7&+ETDjp7~97diiUCvtzxHNnE|F_s@dGcOLeVDC!vgGUy917&Q'
    'e6dfhsm(#P_5xfA;gbTOSb(}`?k|}#xi9|No0HY$II?#NGlv?V#VQY44vkZay{2vW4_N$p60$4bQ79cxLb{hhTB0fwW%wA4u6gNU4?Em$a3}mp&La+{'
    't4u`Ukfc+XU--_l36l~t=<=r06f`p_5`)#j6r=~oV3+Iw-|A96=w}>741-G#Sz11MbM&pQM3eM%FiG~@*OUZj8j0&|ruRgMQp@L0g}}J0O{|(t*8PIv'
    '61Dc+B$KvhlQ-wwO?7s+eYfIt%lBrH8h*Tz$d_j)HVtRWmP8y)TRfNY<#QQL(w$()Eht@OzUhK5#UNE|qgA1SQM3y(^&Rh^%d3=&TEj7i8wk9c=Ek+>'
    'V5Ka=7>P42+F%)&og$(2vQ2nY2!!VT5JI;VM8(60!}%6?INk<TCB>ZM^eR-&W^G|7%j+PDj!*tOo^m?pxxJG7l4oxCJTx<{EDwt{m>J@l!U!qJ;Jnb('
    'QQ2Ts^5u#)&yjQDwGtE7a8bn#oEcPKeIWiFG_U&g`z>*J-sDlG-QU@Kd``T&HR%)P-OG(nOs{2w@y^ZP`hJN=7%|@C>AWPtmCNO2#dcpkOD7bB(X(sx'
    'o&})R@X@MJ^qKkmM?N&Ymdyl4wc}AJ!3R6gXoN(YBU$}}5xKmk9+4=txmkDgzApRA$NA@+M2WmHw&2(y4!n}%U30fWKOlt!E?~7?FYoVHb5OQG`4Q9t'
    'jM~Ri#D9^up6MqnzO4_Pem<qSTdy|f_>Zy33BZVn8kl(UnWkPaCAcl@YP()wRK69ISiXxtQIh=*cjDDrNI#dW%Os16<E0LaEAFbWUnx$K9-{rGyay|t'
    'UqKh~cT~%vY_Oy<&<_OTm%H2duNTPh6^GT#-bs2saDn)KxjVxBZ#jbj3k>-$pKQ~8{+x%DT{uQP4Xvf#gC37$ld@4VG?t<E`Jr%p4TefCH?-YpYHul)'
    'op`|1;q>I-H!$+tNsz-kEAxRB8Pcgzikh#pgI-gLkpD3yVacG9>I_O44f<+#MBh36*2~!**o!&I<5ce}10ELdB_kFRAuE$vZ`Z3GiYyWD5!SLhPR@;a'
    'D7|-W0Wc>3w0knA0dxUyr|v*M>C*tRqk0NM&H{v|fh;l%lj5Gh+z~nSe7(9kGKFj1xCIvDfLhcSot>lsHfI)H@0Tm$55w>&05!_ilQXe_6GmhnR;Ix1'
    '!3=@*Gb}P#5Z42i4PGVQwl$?{j2X=MPUdCD{E4-z#CfzC%9`Z{yMZtc)!zwg7`gig!R<`Mde$tc8o2svaisn05^R9uepT8D(W~2L)u%Jp{Ogzh_Td}S'
    'ZNhZ_`}dFE{`m50gzQ8b-Dx$@3cXpv?ns$cSdgnd`I*g=ZZ5z4@u{``atx#kkUkB>Qj0ZG7$NB>t+vXW0n&O&oqX3vRE2wL{CHp}1KHIfnsDl}Mfs+C'
    'E}I#T>dGiw3*cYh_gFTSndnkyJDdwu0J;p1Ww@UmP8G|_KXt<yL&HsO?o_Gu@c`!=y2DqQw=8X$(}AuG)b({?|9PMzEjOf;A?5vqVb*N1@3EH<=l3{T'
    'a;EhYDb`nWtC`$Pd50ly2l%q#3P{<C=lp>7aPl;n-W&n^+&6!e&fWXR9Tr_6sqE$W2}{E`SKiJ~!rSFTf3v--CZV1H=>pW91iD(mra^xzKxPb7eP1#O'
    '*GF`KmUogDV`7c<2BQ~)1w2C1(sRM-#662VR>c6r)z5?kqxJ4GPY5dJy16+ZXnZQOpuw58Pa$!nj_%6^T@d{X5%sf(A|DNi;f09iSwxW*PddRbLA1{y'
    '`poq2J6{>5SkS%iFZ(x1rRFgg>gaqFmXDMHyJ8_y=kOL|SGjak<mO~zMKS3*<-%#zq-2hQA#6*Hk6?5~v&sKhdCBIOaX?OXY9g|wTrlMh_z1@H933Ng'
    '4R`R8`4okRZ$AI_y4gSO&*6_}Y~N}7Dx*sALdQu|%EAuC)=CoAcuK;e&ham=GlInP0HGToBu<9s#|ue*+cBYNGAsw3I92{34bx(fU>Ch`owG`N63)Q?'
    'fg6@(W9qfd^c@MKw0Cr!<%PW=jbN%%N&%b062moVzY!j4q~EFJHvmg=cpdkEz6k5CN1J9n2Nw#Qdnp`odQ7p(7h+Do=e6kBT-qT8{yXLZ;C|Dj(WYF6'
    '?KCR9fU8<8Fy`1uff3?B7hyUuWj7mT)Oqe+0)Qt88nTbXMkRmA%hmhEc7c7Om|-OMEcqf?iW$E>(@a;6GKK_aEWXOy2K(H&3yqQx2~I*MQPDnH!PQ#V'
    '$*sA-7jhx_h=~UtoCNWr*6{_IIjr-jyNO0cLBUCkA!?eZi+<N%+`~s+^xTd#7)}9)3%JgDnG&^=G!i%gQr!}{@LyE4Bv>SxW=1s4Ow*jjbJf;4j^L&A'
    '3w!QHp{LY&H1sbnh!~>8LSmS_#8SnV6j;)>Q0vuJJqeF8O(_+#-Vn7`o-jp4qh-QJX$BgKzYV;_PKkbqVj>+3C^3;hX?{(oR+U0-u6I`3b9h(zq=Dli'
    'Bf-mNDiSI|T5@#)bHj^##)^4G@>rk8YR=y$*5143>bPEjkM;JZqUv33;Ef-+<p!CIu2)zKen>JhxFV5(thL3;fmN*Qb=W??4uAiTf6oiJ)54yrVil{O'
    'j>E;<mO=I9)DJILMrJ|<xSXjQRQiL|LG%=I+ih84a(`q()#mKo>&=&6KK_Q9F;^*M+L?Tm!oR&aott}mzT58d7?(4sNXn~b*9|GU;HRa5<s9X8^mcb*'
    '9d#vGW?EgMA=AyQLxTnz4RZH6oL;5XxvQmKq}ZY_o@C0HKBjQ_S)*UKNGpi{#Pg9#7LLe%bKd5{=uWE2mEpq7QMnemTIt-KRBd4w^7jPf+F(+RH1ft#'
    'w33b?7`(!+^wnFe<k=r_+V%^sxWUA1t}*Fcj&Q1^URmo)@Uvz!YKAi*OIrg@bmfu!iAUmh^z0M)GLM-5uw>=vbfy9V3<ux!8%|fHZ;CZ3@A@^7r|?NI'
    'il289Z=&P1E)*Efu|1CCwO;3=GfIx?!<;emIki-Yjr+vSHWTB(KzY+$=*pX}gMOsZlq$}q{Z|5H3Q&ALD1EuTB=412rpBE%UC7LwSpDNpzY(xNu)`nU'
    'OL-wWgJ>flwq`LB-uvco3%)<fy$J9jA<x{Cg<YL?d2xbuEfmGcDcI8vV-^nq2vT39)bs=0U@;IJ6Sy7aAwf!#G}(%|UWaS`5`Fd=RaNJP^x{l)At$<V'
    'cr6z~1p#?9)nGo#4mKEX%J^7~&vPN^f}Q>2^+ckAng23@U<{`@{BW7&F{)gU^Rih~eDrztwfIvf8FP5EWMbA7=A1!z2di>QVpO^%^D>XEd?XI;=R3F}'
    'HN)ecHLu`_Kc}N656lL#@~q<t4|_>PW96bqmvp~hmO7U-HL#OCXa&wLH>Oasp2=;5wlF2d$ch*biyR-~VN`8_#T9ET2i^}yiWXQgZpK*5=SNsv!#&38'
    '_sBeday-)@^LU@n+kOw5|9v{5^1X~96d#E24zZ`H)D>2>gt;UpS2Dfn7jLIHwnZ|+Y+Yh1*(SywD%u*9wH)_Eq>&D_$Mu6d4-tH5$}O8NQ*9SeE?q!S'
    'mQ#ENbi{iUrCnV)_9QQp!3_URRkE-nF5logGU&^C85YYh``j>B57S0y@3}lQmw|y1eRG2qp=D-Pqp}~V-~?t7t^+aFQx0ysW^<NkdCKtAQ30pR*)3++'
    '`&WfXB;iW^tn!M^&#bR^w;(rUbx%21qN|zPK4E=jMltWXgpgxGDYsEhtE580)QBt{vNIV(3_H&;mdoQc*<aKOj55LKF$~n>mCyYR(dAMz_&ckmZzHu4'
    'mr=88oVic=1<-y)hKTrxuwcOCAgaNu?S^k3Z2e0N2Bc<SDewj?SFrS-!s4p)8mcQ$?l~y<yR>m;EUAF6;3P+%(%bWb1oznI0%!1J<_L^lj{IO<IV|`i'
    ';zxORe0O}$dr5xJ`I~PuhR)}m_d1kwAO)mJ!5ed{Xp-U7x_iHE-cTdI&He|tIJP{1t2L3xd%-L}HQe8$5GC_njEDgNLcWgA&0)AqPtiA4FQZB3p&*qB'
    'rv?H|{D=N9oU2s)z^wRr-wIaL=hVQIqKf98;s_{<O_hbbX3xyIt9_OzMQ-qIbnOmd_bY&a6`u+qCRSJ!{t6sxYFX}bppM7pJtK+_5PWba%3zb=Yvtfe'
    'R7+q?fuSRb)Q}tBTr)NG=KF7dd?E{`*?s-=`NKaxe*GY0^P_i6Y&B47tQoXxA-DcvRTa#8QtCvRqLhxp!z1nF`FoNZGRNNq#!F_>8D8Y^CjiuePb!S6'
    'zA6_Ju1jToSA|0Em;EM{nLRc_0fMCx)mUG=ObTp`B7HSmLR*G)Z8_UoJ79)+L!3UzA<aSF>;0-+R0;nEIQ@y!jYP!KSHA6k>RA$#hA*#Vq-W9B)WR8R'
    't;r3i9JnO;@)7G+*5qr#BZIlHhc6^B!KFwu#@}0Cf;U@kALS?v#y4;Fd61Hc9hrId%h%sO{_&w%U=_sAA7FBQ|C0LTY#O7*NbIb4Ki(NQFv+@<6UPzY'
    'WMX%mv#RZ0vdl|>_iS7R@=8E#0qJT-n<djmI5WntZ&B9QXTI)KCO(`sC6@M-)oVN!3`!?AZw{GJK@@Kip+XtpiufR%T;mZpdwiv=LfBt5I6?*&%AwGz'
    'nX%G7anVSXa~^DT)^@j^i(jXs*Jani5k6*-L$=_?1BBkV!QV;~6$}*xqK;z?&&QTYCU`qCgApYPiq|iQP$#qe3VlnAvP}_l7ZiAQsg`*BOq<n4UYRT3'
    's`Ima894%whv$*8I0CGsmX5pn-R-bQ>NXjCO``lhQ;w=?^2D@a19u_a!Uh>!!2*DZvb@VRVtfMOF?g1K@qhAfmO0}lh@yVk4O%m-kWQ6M&G`LUDeLJv'
    'ZAa?uOx5n%=JNF}|2~Q`0vUX87Thd}>1&JoTixT|=B?PZr7t$Rc6eyQ4Z_=gv)Y3vD!<P);UQK3rr)!4K=q>Vm(?V7t+a?LrURc`=<<l>aT!viIBDsU'
    '6z2sSm#KP_=ykPwK-p7K-}G&X2gOje@~yXbHw~3<pju~}2^~F~=}8|vT}U{)FJT5zT{$Z`=Sy@#x(HsM;UGd{M!LFsEEBJKoLH`+PF^M3$a0~G`2jG-'
    'Law^D#w%6Z0+nf*bX~I_iMwyZLq}klR20GVvtY^7<uM}^(-sta0wouB)|~DQj~$`q;!H$yKR`=&m9P@uRD~6~g4O*1tGq5AIDpb3er^5eRh34(sUcRS'
    'PLm-+;<|)3#5b9C3J$ka&6*Nm>0>6H-1Nj}y8TtoQNhTz1z@B||11V`9?e4-(a7ae{1Xtu6g3NCt5yqt8ibX8XA$xlfN!Hajlgvr^AJ+&97EL4-2{C0'
    '1kv!Y5m1h!;9G4^1JxDlb2ntQDX-Q{aU^R@k=1z!na@56FQ-&aD&V7BbYiqB+B*G-G!Em9u29YLc;W&j09rrtsY}H1`*d20&^frlp|EEA?rv2F?G(}x'
    'a17(_KmNUGu`Oo6M$b?~@Q@4teA3g%4{!zN<U^eXDqZrt8YcHYLheK(KD1Iw`8;we)cG;Au5LO$2c1j%dF7mV<((dHF0swBI6-w`J_X3LZWa>z!-j}P'
    '^jo}li6sv<JdC=4XTIPzfFmF4xgo23S|-KGEO1SztZl&$L85l7cX$GyGQ%`2uTuW3f(>B|n^n;eHogcpA6Nt6q-RrrrHT`PDZcw-;Bwv+V1yY0b450l'
    'w6Rp+Dd1}5Y^a-$7=kNo^i@aLnTQ=rdezIYvn1=Yz#<*uyFdrgJxAbr3>tyuj1A1I%`%&BcuM+Q6*A}}9OGtv>X0PW!B^bnsNP*+>RIluj~mVx<|imS'
    '6~WZ@SOQP|&v?_UG%##3Ufp35WQ8%CGu!;jm*2i|4fn6VefgI!-@x$y^!tDP?ZYo$zQV>tSVMG?t`Dq8F|z_QfzZq%EGrO9_FTLQ4?;1A!xkwRl3_q0'
    'Cfq&sYxkrg`s3RqDL`ZM7q^@#=dogx<<V5k1;5^{k-LkK3VS-@Xn0Lv_bXS<*ily!3)^Aa_=XI$o5%B*nnJ;F)pEXbQ|FWcj2nsIZQmc)t6hJ$92W8<'
    'AXAKKoqdC>y_05_Ru5W<A&pw-8dX?weIF8y;~$Acc)i5};{6hxMCz)-xrjYD$KOBw^5K_{SRg3zIpwTXHnc_;flA*Nc@@#)<2s~^WN02Put8_T^=4<e'
    '`0Q14gUrd=EklJ4cuHAOT8@7XU4T5cDYd#2*(5fqI1@5#tsuWvTW+)-&@LrnaIe9yLMqXaJUI;B#I7*A!?)@;OA;q6)Y6*`3^;GoMxn7Q8|gThADtqb'
    'X)cr8QJO;4wv!rqC~67Ak}j^-0qjGZ)i%3D$XVXbu~cKLfTW(Wita`^<&!@PJ75I!Ikw`RVy|+T_9~a->@i1k^LsW(1!4@W1dc6IFRFa7SQ-eCBOnbi'
    'R{HuZC%bR43-bW}EGf1e(8bS%p!)OYMz^2;C|5LKK`Hnk-j)uWDrLz9UszXw(zndprJFaGO@l)chQ}kTQ>9JAaRW|n7NDp(Dsn87icO)QYlje}Yr=6x'
    '`XmJN&5egkw*d?{&kPqiw=saO9>G2XAUb6ege)J32<|zA#IrjdpN|k6@1Gf;WGTnPT~~&OXNMo=hUYE<9Ns-M9H!G>k4S$tg~(p6g1x4F2BMJNjSxgr'
    'Gb(OA;vTeWd|5)%!J-|Rwz`?iSK1=u>c}yk<|bFhGtbS)`|aH}lP9rH+20)J0AqPS=D9hV{o{l=0TOQlf$R(K?Bvo*<bnzAxrtX)3==@4{Tu}uvVmgQ'
    'UA~03dq;6ymOSUar|klbmY+ZT{lm|AAU2;rJb|I`w_jhEO^Fc|v2NyVu*;=y53sveT|2q|OM2OA3>0Q$rclP{eu$CU*lG|wS%Rd`+SvXGBztce1Q5V3'
    'H*INS<KBuBRLo+w+#S(pNQ8&@&-EpMocjv5!<G%B#p5PRh<T0jl+JmtDiNZZyPwP$C0igirLY%Z^-}%}Z27yS(E?ChXT}(AjFIJmZ0e_-gmT9SlV{02'
    'f~eD{N;ARkxWuwKT(21GKfZzU6wcI}?K&l(?|zCSC?#4pNsxUC^^OwR<;4i%7|NEQNu7L1WZPu-9gxY2w<Ne3WycV{d?XVl5IYP7ZdTdaNFd|4lj?C8'
    'q&p45`%eN1W00<VXfiXF*lq@Mcbi`?yr<QII%6|NH^R_uxe@M`?yr{IEzR}<kJGY6fUykYY;;z-t2kFn4&6aG9ji%doBJCWFhp=|-t`cHJDd*O+4NY2'
    'QAtMV#<lHubl&7(v|?r9T$c$SetTHm4dz2ai6a!9k2pl3Bsa5grCy2QXx)uQTj$H@M7D@;{Swc#rhJV25O(|uY}WBc0FFTLRIaX_$A-~Z*OiNF-lOm9'
    '12NS>@n|==b=}%7hbUs?>Si$ehYBvzoHix*+V#lZM}XFVspeW2yK;n4*znYu!^QiX{)!X9Ll~IGV)k+2Muw5H{|c`tRgQ5`j=RbmhG#90;YcN-s{!@6'
    '=SH#?9qvDaXK|x@fTJ@va;A><AK3?!$1t6h)@A3_3oV18L71_ui$xCmls!8xah;<0oWhQ}ZQI$`Qx(W!^%%aQ<6Wx$qr#8X-OUWJw%O)j-(d=zEuP#c'
    'Hb%(l3PiAT2>BpA1&|L11aNZzdlq4zADwkoakQTsorz24hPy7u_FCg*1@(tA^QklL%;qdIYL0ebV3#T4V&HlZ*hbfxUAWXVs`kJC@lA+F%J{sA;doz-'
    'Pu!PS7`j-2&F5*n(ggs_aug9s)$ClZt8%>_%8}6iA=WR2lX_B?1B=TK2`jvFg@=WvXX>(DY#;kvRhILE&IhC=1Rp8KLt$Slg@;ISL9dJ|qGZ~n<pFox'
    'U}r|x5ZfExPC~_+NU-RbfJQl=bH58>aKYw%sTuPw)sw*CBi%TtWlr^=`%5%ywj4LPKVtpyltDC&xB01+;o#^U2I(?$SRx2Wac}xvN5QS@C7$a<O4Fm*'
    '@bV`(Hm~>n0*Qzy(QtWO6M<~|_9jgka1QX8LoR^4)-3l*_4^H1>nz{(>)MRMI<GtGCEbvwe6Dc@vj}^B&eOR9UI#MIN3BKX0*=ctrX}al3-(+KNi79m'
    'a9lq^G6g6WP54Xd8pz!wN!&yegi!+sjM+F*?5q{x6UGIyUGLU|q%G!-vH8@V=K>=DLkV@Gv{e0SQ~1_lyTlIvWsa)y&*9Rhv>~uU0@Jwfkn>W6X5_--'
    '!MZJj-FR^N26KZWUs{R~Sdz6KJb2dhs74Sv@cDfgv=FVb2R}zMJRp{ShGsY>n&Fr)V(9Cj=0XY5>4{=f5;ICSdUb;X`Y}ZI=n=xqO}4~unh|0_a7Ywn'
    'BAJe<!~f(BRK$19RUc)ZMWtY71^7dW@dY>6bFf>%9(teh%Rx7b4H3_|kaI_D2vY3Yupx+CovDaGxhLcr!j?3!u1Ccn;2A3-J7{x0IJ7Lwg}1!H!Un3L'
    'tFyBxZ4O4=zBzA}59yDuRQvglZ-R*%k`HMXQ>TIi0+~u_d4r@p%w3G7Ux|U!s3(gW_Poun7b=^cjHwZ`70fQa1XKGbqp>&#;gTI3jO_x>1C<Fy0xCV>'
    'Ubap(o}`M!UKlaSvRTw%foA%GAyDWVyh7EAD7vGtKw*CD*E<w2S#PgasDY4mL9o<vgLPWivXwSl|D^u*>nrS$%N))=kj0I3%(=)En?7nW^Wwx(U(Q%a'
    '4+iG`-~@(-R}gY1W_(X|Q>l!vOxV#l>4=1zpu_%twV|BP(38$vOzN(B%(EXQIGvx~jtk}4C|7#|HYH(ZVZ#ipNaVEKRXBvbUV_WKKXI3F0tY*(7r3r2'
    'JmeZme0SIkliF04MKe^nDVQ#|Ch;tK`?zc`ke_*|$a;pe>71taIe)BIq9!{1z{<q(p=jZeBu_e+@9w068Y@lqixuq4H`^lr8P8ia!PL4d_e?t$N0+J='
    'bzzOG<{u23LdofdK{PS>FiW<??!#?kUo+pFWkt%FyD^!YL1_-ejUhyo^vN)Y86tp}5Xw!&$%-nav14h<cKX7N5}J^f+t=WGaevrhi<3^c1Tu)4PFQ2)'
    ';8d(JUp-C<2N*g>%R)yqQ#4Ej4trSTp2#(s-!dwagTcNry+LeqEQr%rmVokdogj8Ih`Bc!5C=q{jJQU{zvm#Q>(s@zJT|2Sr;m+T7xMgL(AZ8uL(#Yb'
    'P%~(YEL(3aRX%`(;WLn~Oqo%q<xCuBg6L)tL-{<;4-hw73<o{~H~re7s2H_Ob~=2K1uRu1&P93<WrEywwtON#xBeA&AD$7z?kjO(%5Yb1%<H+~-C#K0'
    '2Mn6rN3_C9)W%dsmsTwaZq;G5^f)kN2Bt{Qv-8J0V{+YZCUx;R5DE2KZOipLGhFa82OimOp6;$!Jt7Qlhf%vSS3CcsZR7-QH^udpJ~3H{6hT85a(PG@'
    '3$|PnhWBHDLUGKU3Dl6ke7Id*GM{6GQH7o}TsMM?!g=paDE52uRMIJYT#TF+go<_urA_q(V+ltm^G_<T2O-@nJXFHuHMuBT1>DxeF}UeB$M>SAm2hB8'
    '7915EuWgZ(4^zuD8AD>5jz&9EMvH&ck4b25iN1yJ0)~t*d|N*F)Iq8QV?}<6fqYv7DSW&l9v;<+SHgMl`i-P@Peu~_8`)b+L4`>!3bAs-f1;NXuteP4'
    'A&n$!bG~X679N3B!Ngq6;p*x<*G=PbYb-cDN^~5wGi^1Y*2uhBOI1|Dzu*M>?mWliSWXe!$y#w!nA4P+1n##x;-cGbcK`W5@7T#{5|P|bDmGw^qp%vd'
    'DebEah#FqXIF&0RF!kb|YXi*G@2P!?N}T@oYjYD@m}F)tw+c6hJz%+10B=>#u|h2R#r^TnDD}{x=Xu`T7<X>My-q#qHY%I%7EN0K(J{w)F$(TqTH>B7'
    '>2sjyqAGNMbBVy`xgND#NdmuZCu}x<{z@EDdG0zA7lG?^|MYdrET6eh#H5_cVJEl{cY!1Zd#0{WJoX$0rX%jnB@Q}QBOI%xtA-hC7u1ctOr1uhiSo@Q'
    '`L$h+S!IN%RZc)K2P&3<z88Ndv$`F2$4#z6LAn|mR{<$-iB%5za{_g&kHU5H6ruj`_pe``UN<=Cd<5f#`YzFt7V+u~fX{+I91oNxg2Y2wMqVEuRwswT'
    '+Avb&%RGuN;K*hfJ99pU5?U#5ZFnBawl$QmpiseDrH*+DT(^2iSKhxl9!P&8MG4QFEz5W-$v_f}a!Qvy%97tbqG#Yq_JF)S!(cl(U@f*V@Y_MjM=EnI'
    '07r~?Kf)T3h+Z0M9qbp_^I3!%xL5d&Bna7*8?r=Kb**N%s9KU!EbE*$&vV>4o!7E+T2HkD!P>lYM%K_A)!`);UK9tx_cSzP6BZcj*dgg%k!$r%ELV|E'
    'yK9gR>QLWtD%|(qUt=L`7rRbfr3T*4qlTnHj&Ny0Wmh2c#?gL49sm4xsvaPYTjSt#^mhTQGGrSw!glUhqeCibnUNy5ilitFaj@5<cm>H8VI!6aIe}El'
    '&z}QR<+gNt8cYg^pTm-9nB&O$u)dmQ%?=aog9<?OQvg7pzL|Su6j=b+#R99FpSh(_w3Iw0lg76sbcfsfq&n$Ng9-mTfcYPR(cU`)<}{{jR8p0rl|~;F'
    'r|WWE4tR|6ZgV19M>ylpz!ZXolUGvYK1QEfsd&xmT2PL4<g!^L`^FL1CaK9GB&LgMNUF`T0@;%BA@P!A9kQi79>rYbToxNoj>59noLb3aZD&d?865`W'
    'v3s$mX$<l9wmGX0>mcovtLy$fY*%-f-lT3KCc4k`ojFM(pCi)ihs$!N7O&6k$zn)U5&x+rK{JNA2Z$~^tzI%UwaKP(grg=XY|@bvFDN}J5R%vj&Tr0w'
    'QCW4Avg!^tB7Au5D>WR=Z2J5D=D5mz%Zu#-^Eax6X@kOMN0!ymNk0pKyt5YT#$ENL<IFVF3eC$&A_(C$^4=u4fG46;okZLe7Zuc5ri0{<4~aDh3^dxR'
    'QOUbf*lfQyZ{Dg_z7bM}d1FH>IzH^Q4pLamQ%;K!voh6YzoiDV|F<2(iRI#%4CBwiP3}&#qdaRU{y8{3h7(>@fb)ME&Wzz?BLukcr{U}vPPRhNdA2_X'
    '=f-d{Z2&I*X}B<k6W%x2Y5)Agb`^^<HOe#{QQeF#Qq*Z!`|{6SB#z*^w!8w)%DoJZH?*1Q-3f;V0P24Na6G{~T|Kb*CjiG&yr?XCmht?vfa6JC^cDfo'
    '{S$!WY2LZArlbES0LPQ2^JP`X@J|5xF<?>t^0E6T0LN>m>&nWG@t*)pM?$hr@YHEbTJuC}YLO@50;R(^q=TJupP#U6PV(950xC)w4fgH2`3szE)U6ut'
    'V!B_I>ZF;<Z$NjDE8S*QWlCWHy$Z0fLUvpk8nxZBTno~woucZ^o%cV!Suh0nt8;@D=L5(UK{`iba*^Qep|K7Xlc_09p^0Z+)mR4aPFm&)zs&gVL{0t$'
    '_vaC_El;fHF*%_7qt;}G3<#Oa6_J`~9`aPr_XMv)JS!8(wDWjm&HorqwQ_s(7X3-M?3=v+lgq+|^xkA?QZZ_6_A<^=sdFVn-IFdgg;rjrhzbH$_*xmT'
    '3vJ<`tr*qzJrRgk#WGl4q|dV*854t<vr*Oz*S@Daks^)d!~MQ5D)o2Czp()w8?u+9>*@&^vENQ=fm+s#O5{<@-c=LbpkGSqq8h-qV&WM-Y)%4jVP%G~'
    '1x6w7e(Ju!@Jh`)1Uo_^E}ruZxH6QY$!9Rs#qeqfYm3vN_^ZAImRt@6)^?IG@SL*(7sxtMMu?GnY4-0(HzN0B->myh^Q6JPKH`Y;n5)+lhBm@df*4V#'
    'h*CWmk5@&P3~<c3G1YWJr$DhJQzD)nG(&UKY2NKq?e1u_=3*iqJ;PB^@!#=%RpYUp)@0IrojQXCN0`WRA2R19Nq+is95_HSt>~Mpj#DHRb^qJdRh?W`'
    'cCjg=!31ry>>4@dvpy}#)^KD^0!WWJ{*~z2;if`8aWoaeM-eAT2`Quj!Ja)NPM;UA90X7j#d|nMsDkawnj?xp=O;_16oRLv{FD~ndA}b3*#S^^1}H7P'
    'wDM{kt7FtkD!5~0?W7#fWFa@^VwDw3ZPXZx&O$++JM)tO6waB|9k$z|;y1j5X0a%)C$%80Zf@@dO~X&H8oKZ=o6>$;iVW-(m?+k5#^ZEZ2UOU?<4V5Q'
    'UWu%?3^jQFN@Z=Y2o{+XBT2AH_I`<`oeii{p41Qz&VF(9ZG5f~yG-fK34I~F1>Bq+<EtUKZ;R2WTz48A{+xxl;ZS;bJpZfeJZxsJZQ}@1Sc<W9r#-35'
    'k_6D_1h6a77r2twroI9mSwe>CoOE&X2dY-k&1%1HsCN62W}3tr0n^!n>0zU54OiPU#cRu2q0Dm0qSvX+aK&6^SP$`_n$_bu4lGF6{L_j*h=pTCK%BDZ'
    'XMg|j_2+;6c>dw*m-D}W_|O0E`^VqD$S0vxX8;tjLRew&R5f9<`77qB+R^yfb!EIM<GGM?GTaP?M@&QHoii0Pe)#(7j}OiI`{z%;fBE*Ix&HJoUq3Ye'
    '`9J^u$ETk^G?%}9_#OTWl`|_KhwDh#3cVgn^%jZUWKq@M{O0i-M~~s0FL35GTsFd|h;D!g^WEiiix%tgvM;93w(yeXF-CIu$~!lwVf1GEh{X+W$JfW9'
    '6&59-#dPb~h<N+y52P)*VcL@2*N>k+egE8CfBfy!zcq{Rpa1^bM<VfUe*W@%v-tA$k59jUXm-DS_~YY;Up_RxZT|l0<G+3QHoOYFwzgcw%-r!!Pco1B'
    'en-QlDg=kybdz+4sNeVTx8FYCEiXQP{RdpT`5(>l+aF)?Z@;0MCg0GjA*AWD+AT~~9ZqRp@3%199|o`Sa6HH2iyAE$FSHPWbuUWlfgUgOGG(Iq_4XZA'
    'Xb3QT97GwyT9wm_I(0fHR`@OUW|X_cZ++XGC2+RW5zs1M17b8wiiABYO(hzPp@IPxfE1b?EZ$fh+IZf|CWCybD&~594oJt`aG9l_Fy-@*qtLNE*VS;$'
    'owhb}x5FN-+Ebt4mgGR;|GIxn<m-Ff1MmCM+y=mN+l?+!G>zdHFsClD>wkj(++%&$45%v~+qDKPaVw!(I0)3tWR^aN{MiTA(=axn>^R3m*Kt^FA^g}a'
    '2kd8W{_T-^me>g%O;i;^g^9H`DhlOgXYA=0^2l`Us}28SjH}}70*sqi;2<VBBV0F&2M!!}JY#t-EN<aemvjgo*EbZ%uX_-%ajf-ap<?Tm$||J{{G8>n'
    'TfOZUw?|oHh^N|O{OV||RyGO1p}Wq4XhKXXh6~<JHd+Q<1G(`8K>9JPD?JBJ&A|=|US}YkKpJbD*1l-LwaO0aZEpMB?y<RC?eA}|`bXkCZ!VXYw52T8'
    'U@}vKO$c(Vq6w)`Z784WaKOgN*Qd(7QRvLHn4L%MqiX^XMDzev+opaS14kKUL5LKBTJ5m9h55p4i%B3U5Mds+4CFAz1KWpm=h`l-W6!HmXsV*{&CM;2'
    '$}mK=>a7SGgB27jyJP!?>^+GmOk0zXB7U=Jz=de#j7D|{<o2vsCXA_NL(#QFn%D;{A>{@hqm}GJ9eLKUReCWMX`{nhF`^T|-p}QF2M^W4wUe7%2PTX-'
    '+o<e5gw?Gsu<A7I0M7}YmTyVc3_PeSItP;Np6<uwg=VrgH_%3WTt6U;><u0Jn_f<TXDAB*Tgq<L?^e43b)3O;<D?i5Wz3Va>E$@<LbT3#Rhr^BEp?CT'
    '3w@>M)S#?SUd#LA=8>IOegX2GJ1p_TPnR3?BG-hxA+BDmx#LpSTr-MX7N37BqdTJClh+ryOlv9(wLDzoU2P8-{l<o<fw7Zu#0iK@F?9;3E=ie$ymfm5'
    'UDUK9ORk<qAFsd@Hym4`dXG$favHMq=j%8Ei{3kI(iRJ)<=DiiaPjg8yIoetg#~NOdFDZg3QLwcQ&bN<6AfmaT<Mz@wOtZ%UhS{<k4rp&{F8m{UZqKw'
    'V^oK?M_AYAM8`b5j%ZMLc9&jy9o;^EoxG`G$|IMIFn=EDH;Qh2=F(4rEk0S^Ar%Y$<^5*y4gleDNL~)t68`cC>jw6JtVYZwN}M;PIB)ierG757Wf)lU'
    ';C+aegJa_Sie|C$ebMKIJ%?OixQX`HuO#F}jtF1~sIkrmZDN!dvI}z^lgt>%t~yEMl`7B77!9_GmcV90#s@7Oi_eA^U@!mcaYb#~{iYxDK`<ad#FA{~'
    'YPs2s61IwaFw13!EC!z7NIEZYA>oT~&~1<}Vl>v?StYqp$M(|mjjPM6+_(xOd%<#@v*0DLAd=;@PA-SLW(ATI{A_Ye^uaY}i**#M4U6?U+h&$-iWgC9'
    'W2ymUd-$=K$djZ?MZ2%)Q01@@_^H)3;Goxdxa6K>?rP4Fy~XWzw*w!1{|-sa&QO;J&HL4YLtC(GL@nLi42RV%k_+vakE&8KVL1bqIJrk#tI4&+8oe7>'
    'wcW0QbJ+GSMG2P4yd>@@FHvqqHZX6jx2DZnfId+<AaN;jHRIsBW4~q(kA8TmbQz36&lK<BS|lpy%Dk$}^|4l8DD4<Ih^pEL*UBpgyGK%d89S+Za9cCY'
    'vqX;GAMdwF5_v$zN!E1hS)F*Zy?Q+?{0UmCF>WaJTih+$-yd=E-!IV*9S{&;b;X1`A>qopfVZe={#GA>!qM1Gd_}||2h>;^a9b{jYUqYKRQB|%HHLZ{'
    'ZTPggRSvdiw^A~7z@hReH_}OhhyU0guwWN%kG!&K4+8_N?D!KSlwQ2{#%Ou*0S{QEifLjwKatN^u2P-GFK-4G&{M#zF_@swjm#HZKzPS0mhCZA1%MPn'
    '!p??aAz(3?kq30OUOZyIb;g9H1pZJg#Zk5`o%~3{u}G!aiD)2i)>s0KVM}L939*mQxO-G3i6(f@4@#8L+<bvBl>2}ax}v3xIweD6rC+V8R_PRk;`E}V'
    'dDwEk>QQ>O7bKDpX_-SDO%<AaC#;n-jeNrKFQ_2WOwAT5ybhSRa#GVg^NlOxsg{{UE#BiL_pKmGiN#VZRpT6uWkYWzDOc}9ZFAn7iBxAj-YT7fvBas6'
    'q}Fo~_jY-S@#(*?KP&ClIHr14h@o30hb9X^ohpAl<zO^wj3`PWo>}rvA!0Z9VsH>sY)98GN;}sosb=n2ZHR*A*lz}}pW3(2xQx9@kn_7p$BKjfL0yhf'
    'v+NFri@Fi&dbx$&^6F*C%Mr4cSK4t(uSUM=*BdTk_%ZOh0(KpCo~TkHae|GSYNvlte?akSaLHR)a7ayYLghUvV;yK3-SxQDbd%C-hE&+KEV}k#64^_T'
    '(_InOaEL*o@nQGA#5KCEh9@c!B?$<!8B<vNmybVx`TF7KPyhM9es8Ye&&}nRCkllt6EfBJ$KRVjzWfe<*nj!<vAOy3+b^%j=pyCUMmMbHFQWHCEW{ow'
    '?B1q^fjf*sDARJXc-DU0tnThsNvtfTAyiaDDBZ>qYwS*%qTGcBXz!Lp=8{>jM_@XXLZI$sY@HfG>Y#r~l7tapIsFuiBUkTh{G1asIq(F{l#()KPolxu'
    '_c((#`Uu(<r8pKnn0;9Mly85KW+Z$&!1vl~QSGuER+zinI!Hc(Jh*d7V+m$nJm`r$H*uB=usZLjfc0`w*t?{Sl%m9<5}p{Vs;<KgvVUJNs|!+*XQdD<'
    'afu~Z%eY#J>J(N`xIeWckC7zb14#Z)@4^Mmx6Z=x=0&BCcb}Y)`bOJA-Xd}$>Mxio`|xl0SN{Y4`1^nT_{SfgaDVs!Pi{vkDovC)$U>h<@Y_z(;_1d7'
    'Nvab!WUyEltIJiBiWIr&@nWFulubXUM^_LPO{DkDg-JDNRa%nCA1~5%BzT?CyRvFM(OtAfqsPQ&!$M{hSt5iEqaIe$YUJUK0&;M;{Q4E|flF{&S8kz0'
    '{jkUdeyM`Oq*q8ilaCFXq&Q27;sPe=JL(_n_lGe`SA?)7nLdv~uly%;*fR#o<Y9G`bV{VqI~HwRuqDPX;|8<yTNLG>)c(ZWCf~q;6)N)tbNd>#Bl790'
    '`1U4kE9nzr1_Jx-uBvZvXp-B;&g|n)DzxHKY9ydArYN;yL|mVas;3}&9YcqmL5I!+p(ny-{KTQj<~5ur3u1Ii-*Rz_L~<S*um>1HbwaJ6yk_btxQhd;'
    'GA@Z~7Hgb|`SOsX=?{ffkgD;W@o{)_AvlRWB^9Q^*9T)!GI)hig|d(GxULl=L>oHjHpl<`zn4VG&v&#XMZxB7i#;*?Arr;+1^Fu~KA86B&#xa8-J9Xw'
    'srn+FE@Xm=$v}|2t-m`i7fTk(AvzqmJ9iFaY-W&_@kJq8F)0(qbq|tH!FTrWnm#^Zc-V7CZyqqD*hO(<d<MgeF!UHhB!dA%&12Z+$yX3&3?b5w%4H(f'
    'F@zgJ*fE5N{{n=Whww;JR3Jtl#H0WjK_ndkq`^i5xp3o_(({fK831qnS2&5!)de0&BS@#q`%S_C{lLqcDQWN83RtOgKL7^v(BR$`P4V4q)p(v-&ezCQ'
    'cBnhguX-FM-kWh2?QKy>SCOmPSOYz~%LaMB-!WOlai`2Vj>r(G<)LlNK`oVZ#_<tJ&4MnToX@^#6vv>5GXlr;=lw>Y9QHT|X|1)kvnPZ~Qq4Z?bUrJU'
    '<o4^(mMLywVj|IN_Hu(YS_^#_t4i0rIJnK9@wWfSr@*T*ROW0)D0>o08umtLc$3Vg>Fpvzqu+pOTM5~Bjsgx$(wvx`@l)I^H@TCrzYxaeZ1}5C04-HJ'
    'nj(Wo=h7&G!6mo3d?jyT8>6GGkgz$il19<0ZX8dgo;wsPxmpsP+GsLY^<Y?PVx?oTfO~N%x~|IFUk?F>xCMi%Rdnb|YAL)ZzbO{#xB{Z7oRkD_Rh7Bw'
    'PL>?zL-_)4{bEd+Il7FtYJ3&hDH7ArpSnurzfZ2D$5%4tb&X3Xsmz5^0%)}ywrjb>#qIV@f0&}1F?#u+lS;ai&v63VkFh1k3E1w$os&cl1Nt#wR{+6V'
    'f-_k@Ef^u)kNu^(rJP(QtWB!EWK<D+Ffnc|3)dFM4qgD*q6boF8)vzcAJz2fVt8ZYd_bL@wEuvoL#K7FoxWSrI!V?VZ1`Y7H}5Zzqq)B@#BiOEyh3{0'
    '*h*YpDXJYA?GA??SKJ*}q#ddzTn|R<lj^ALTpO_FUyYEH!WSXWMNTnYLxEQ#!G0YD5e?xnVqv~e97!D^%2vU985%2R3d^8sB#c!h#@6J`&OCKKOvX>G'
    '^!G-$`I)URxE>j#qjAMc4yI6i<U;;@WUawH#v*rzegGDTF(1$wlGf-cs&5ObnxZcDura;}y1a-GvFp=u>W&Utwzt+t3`N6CI_y@9$EGQ)Q;Xx}YPlzN'
    '>E`lQS|KzBXE8uvWl3gMHonQNB(C}3%F5GpT)AR#cY#V6Taj!It*e_p@th(ZEJ{T+*NPfdUV~$(usK6=_sA_ob+p1t4F~Bn-xy}{<4?Lh*Ts28iv|e<'
    'GP%38hI~u7jm>sL{G()b+>*Q7KVbcQJvJ&liwfawZpeb`F5i$Znd1_PD(SJXNx_ro%L+`FT7UA+0`~#oDZL+v&zK~~NDtF0HNptFmIyh=NjGG3dA!;V'
    '^=tCv2}ESm_}30IWZM~6RiWgido{R-vQ@m=k+C3bIXx-RY~QhiioIQZNxLeKM92BwmDNMoU~*L(z`AU?;V!vI3d`^<F-JkLe`yj$0p_67)dgA#<}$Pu'
    'Qd597r=jmx;4H#_3*=*ZhyrZ~&>CKNfew0R1Eiop77te8$-Q>fC)84v0f69>DU`L4IhFQ-HR@O2_CNK+W<hrL)poA;4YPQ<m18f01&4=1Y5GuA{|QQ@'
    'Q$1&4oJ}c0tPnb`ZyNKu1*<R7`r|%Jjl7*L8w^|l`=gAK(hR2Y+3mNRZ8@v?PwVZ?Dv1P0;Q>}9rsaX{sNkOrGJOf(szQ@O6&5eXCx#@v2<y4CcbBgf'
    'XRoD8=8`*>{`9r7csE|ex6_Mw*riKT8118f{vslNz|Uf5FXKXa7Wab7%)d6iLN~dB8Yp)u$%-A0$fQ>oPD?WHehcpfe5T`FVoZO`|4%9^j)pGFGtKMa'
    'OjAhHQe7|psq1mwQI<I<Tt-IKmo)9@k}3s>>5{fBK1(|Kc6;22P^bk@X|c-q#SyIyESWK8(p9G74($@*%57DnGTDnCz{Wc!KvxP27!Z|33SStNbuErm'
    '(7|F)i`8M3>eu(n-H}8u2+DMWSL`g<ozYZGLWTswkjiwsCcR;dU~G0P43_SD5~|>8{@o#oR*;ZbBK|@!oIRKdFPEVrfe>_+j;<Nb5;xg~^5oKANQ7EJ'
    '_XUOjm35ZjC~E62&-?!Ve!DG&CuB1N%unynTPZg-Xn=OQ#I$tEbm6Qratx#}i?nS?784oFobSM?msC`!%_1f`CDSZhNU6<{?INiKMxf3i%VUlUpbSb0'
    'wkNzaywY+?u&JJ30_TD!MB2f%-3##QNCP3}!^B{*tyW3PJy8YWA6IwRTxhdgFBfF)gHzo$Kf(LwdLC@D6A<UXZUp-;mO!+bEU7`GIYy!Lc@D;i(cYRM'
    'O1zrIy<qRyfk%iFrXoJfIQ20Jr{q-+50WR6ao%7FcIjg2IJ`PK$1Bx8F4)Xfz-5N=fi7OBao}aPzDj<1mSm|KvP*kTRjER)**E*(&n1W~Sd#wII+|y&'
    '$lI=@lq*axm~fOk)w!;GQJMUw#TVJ{${3wd*n`q$F2>Gw+>kWX>%dpf-O1cTv%liih%H53^?ehZ%F=ZJr`sRGt?qIa9->L%jOm<5tJtPWx-JX#Zu*<m'
    'VvUV;uq)6X(g=%+P1a|7hCi7drtFvza<e*MU1QGQO*6tL(|B^J2!@n$sq=!8?8;77=(7uy_7E}mtkxJgXcNCHi#ElYIlIUyf;%)KzK#~P)|ppK{BQb{'
    'yb#2ZukwD`s@0mEWejD#61X4@+g*s0qGzgcULEj5!nQ#~tYG9~fwLza%?#tHfLs}QP2ex*3^E5z^EZ$_n!g^27JL{)lzsuz6=(zhQPLz1ClVz%rA0k$'
    'A7@zJQFu};&MNf%v?>&x<ge|w=<Ox4KJeGCa&0#l8E&&_U{%?wQk*Rg-KjBhC6{f<r%1?G0$%NC<qFvM!dDRJwm`#d^B8#VN&|TJq6Y9vhJtH&Lt9`w'
    '0L6YUA)N?JZZn$2W&S>n4#XVMmY9`5cBR-z#|dN#XbY|^v5RrcoZ-95c+Hnkl7I+W={n>x3d0Yg8A|f1f+M^v9jC-zh-1b$+7+BXaX)s9W6N3AnU~|_'
    'lflL;kZoNEI<H*}`os3|n2)@-6eQiO7BD+ue?H#eyulS5o}wUi#8)g0*RAD-`>|6yIHEPjNOZ_IgpHK2EPSd*Wlj+4D#0A;UKBwRW#V{ia|*d!Q6QiM'
    'i#j_KP=2n#b6UgZxV^WN{tFd4YhPN!CUhpGS%wo8Ib8;+o&eVzU8|D-T<F}1Yb7E1ixhGu2AT^N15Y|huo`x`T^TFzhS!L>T?Mb<=C48bn%d(fsGT{%'
    'J2Q-JQ@!cn8qsv+IlP%IotwTgs#O_07R;Li>`QnJfaGUu<x*&{s$3Y28Oy*-$hjW(J(ZzI*+Eq%bM}u9Uq5~M)_nf*%g5ioe0!C+??Wr;A_s09mUPGB'
    '09$Hw63ql=8Cf>!HOqoMG=WIkm&gDQn-&IexCK0kVmEJ=*VyPoMesg~)9A2|)F4-^7;0VV=*WTaL>4sDc>!m_kq8GbGVv*%XrY9l9Tuav#<fC=&~oHp'
    'WCH_SqVA!f>OpFtW%I^RvmzSoEMXgMsQKeAR|_z`&|C$h2}}y=Fhc7T-fpLzB&-D}t5xIfR`V+|Rd2yS2*U8KU^wlOk<^r;NLFjz!RCXfXr5^p3BzRb'
    'o8xjlpL9*8LloqQYO<D_99WRAu0S7$1||-T%p&gseoYk%X)ddGjT<K>trFNjL8m;-LqQAFZtbL=#d29A>#!vew!$vYMR6EgS3?BfsuI|y2V;(F(Q|vZ'
    'y4&W$1_UG1gwWC!aE`EnNL^|Hq;vtCYDI{U0-h&sF_#eIO@<MM6U~y;4SoVc`o9{2u|t?Haf{HilC$yRXT8PH8Tuh;qF*%k{r(NkY<}uEd1|YSRMgrw'
    '+8Kd7p%&burzuU=c;DxbM%xjdZ7V$U9A4UDBsQ4<-4M{&Sk7tN{%GbTI{;s*Ou+!hbZ51Ztg~&cnG(BC8;uq$@C`oZ1zU)M2)iHY<=C<l_*vRyYRebx'
    '6Al`|3h9SP;Mll=G*xJ4@BTENE_hu>%{N-;J-C<jWzRZU(lu;^T}uka;1#w%fIWrPVINx;rPUFPMdzNy>JM{W^@0bkPOC5nAgT=<SYG990>%GU%eySu'
    '(WuJ8!0-SpG_AFedD-q6N(5?&#rQ^JHi;EM{5RKm6`YS=cDv5hlnv?8N)>EA)PMf++n28&xrsF~;Q0cI?VRudx)KDXJ6!DHh4sjWP``JQUQu+c!_r_Y'
    'Mc!Iv(=x?+i34NXD9XEJdM}kqTz4cTy1a&W(G+XALApjGj=rM`3F6)^9bZM9^dw2g6Utx=nL1E`<5dN#Fk)CZ>AS8j01_%AD)*x9Q%n;)z6ovFn;FC='
    'ZX?cO`C{;2EptSQGgvvhO-Zd`I!jYIu#<2OLxWS$N&zA_)>Ck%|10sHW0NS{FDV+ELWoH-;+)wi&rz)DRW3J1YMA0^cw|<=$fk5EI8Ie_g)a<50>KHl'
    '!S~;PA$EgqRf=vsB}%pE-otXHcam#v_e)mx!~ZJQ2l>aGHHI6!enaCEpvJ^73K|NF5{scv3$o|D4ScIA3`=CVlTeK>;M#J5c?JaH6c8$pT;VZ;oevmq'
    'BWh{1LLUugna>{&%zChrFk_Z0=qRGYUFQ33CpD-Z&O8j(CrTA8ZVCfx?;Ns}M9LwN6(24RtD8I8%G~$d$CG>OLh0D!XUCfwe4s~eD7fHrFMlCo%a8;v'
    'it}o#J1r%ati^f5Op$VqkSBpZ>t!HVArWhSa`IjWE??ee?*xp=L?nvF$lZVZJFIJseOi>7YbELy6q+-RS*7V157v(I-JH>W<eYce+5_i`mF~nzj&?$y'
    'eYw47N~CITopH8=zy7NsA{ffxwv;R+<(4Wr8Fql)^*5`#oarRfQ3q=M^x6zsKe3wv1L0BjD8<d~W5XwVMX2(-A<1FYJ$0|$6H-nL!9sds25X-IqgwC9'
    'Me*=<?-nSFtn`y@2e3Y%?^$?V+c|}0xGBE93*r%@8`N-v%@oR1cc^Ee3V@pf&{|wa#Ubz?+>7?8OAVI$7|Cc=-HQJ)Qhq%HBxA51ITX<es&qSb#NMy='
    'eR5?-<_&mv^o}VaBoiI(<?6TugChS5&nKkflo&F&lA3wWuvk-)u5~2(db?U<m+sZ-u%JTZGxZnUKB>Qym})>37pzzldf~+8{{3IR|F@4{zcd@LH~x;z'
    'L7Pv%e`@Z(|H~IVEXJd=#SL3oJLp;9`bjGCUSm#&o$vc2$vBdjmmgu}j<|e?7pzh`I>M0SC3M{8DNUZ(UZ|3i9|Iu~5m*7(u%KTDBw^2je!CS`|1ini'
    'Pv<a47h_bZDpnlCdlt1a1SZIvf{f(6N>{x>ETxOD?S7BV!jwlZxsS{3`z<E~x9?UPv{8v_mqn&+yC~2$z(1TfCHN34KI~gm9-{Gy6`OP3GCBlIRl8!`'
    'MY2;CR8z)@55!5oC*?C(p{?ZMk}S|D`f`Q;=&@jb%~1*0X#IqY!(6DSB*o5JWLkHna)F(&bdXC!QO(+PHC|k-B$mVtU0`MNaZ6I=W&f}%X@VPa>fTfS'
    '<o_>lk1hSu+2>hkwWeImd_L@T)U=5(%WNvoO1)@OamdIvm6CGL;3#sL&WVePIPhdVxfB-*KJnbN!g%YM@pp$zq^J0tpVVY@*;K7`TFycomJeGQjmfZJ'
    'wWwT!ku_7y##Gs=!C&T~M9!mP+FZG!#m?l<SNkoz;s%-C!N1;6L0L*v86C>p#S&EuwqbC2cCf<n1}%Jesc)94v@)x?ukMzl@U);RgeTwJ_;3#6E}cLI'
    '3<IiQeG${NC&e`VgG*YP7QR0iR7ao0G+}>vHNxylOqN@kXWrPM_=af?bPs-Ti3eEA{3e2n<r1Ao)_IAgVlupDps5LQgoH*&O|T%QGx%X7H6N-9?fNGB'
    '%jIHug|~K*tYuVYKH%7ko6f)@h`q2ZbSN1`Oo=knSj1uyvb@CAAd!g$;4iqemze%x<1UZtx&Z!>(OSMxpg;Z-UN?uhVVPpjK!W#*jVPJk-r9*6PL`z_'
    'vLn<2RihP5Cn#<cQc{5zf`y0-iy~p?CpPp%k}x%mqylu@%~1NP&=z;zAWyMzrWn#~sLc91%EaV+JeA!jB*x6+vFI9h@K)Fmv`%f<Gu0!URGpDHsR*dj'
    'O;SOL8eK)HLQC(v7hPo(oupSQ@Du!8&5lH)afm-ay~5<w`V_1m!Rk_4V_Nm+Urx$rSC`P^_b+Vb6a#zFCAhPNu@5!28DSeC-qa@Xn9pFt;d!|{U}q~a'
    '8kMRbGrm{b%2#yuWTg>rZ&h)T>=MC|dP0=kOJG5r%ilhI`{U!+=IUc}_~Yw`Z{M2!`yXFE{PM3KzkT}0mv4A16bOYZy((5)?98EbM|zT*;{mo8Tm?!$'
    'D3!KXjpE&5=PzLWz&2*rKG!{#wvK$AP|BpaJac?J$X*10d%L0o=$1U{_|T=x;YIvgU#{Y+t6*(MuJGgX8VY?Pw-EJW$c1EMkfIy#Jh~@5I$1=d6m52)'
    'Qv{sJj5TtsRmRlV<fAlU3cd*wWMxNWyiOH7`wU)XLOuhE(G4&Zq`tv2yV;<u>i}l;>iW9S=`yvH(_<8AbxeGDvnWYrvcSlK=?u2tsMb|8bG`*-NMW?N'
    'IBcm6eAtm9hfL=w9pF1(LhbhXuzeD^g;jJvFy<m`70MSOv*1i3-z$w@?8!*K?U{XQ&~jyw<~%(j?7`?Hu>f9jl4?VquFu9)Pu;f|oy~)>B>{XBTM8Ik'
    'J8P&f_K&zU2LAHyVrr$g<0CZ3bs|zcV?B<~v(8W%QR=q6*r5moK72#vrquheTEf@arjL|yq9e_bu1KS>`teLbFE^`bU8y<q?iKL(^g?sxIEBg?igd<Y'
    'N-23SLvo)5#!Vw_nTvn@^79wu)NlIV{{HFrFU{`5*ALC&!?&N2WdDhIFJ!!Qv}>y?p<!XG9TmquzG-vIU?3$_VuzMOq)stk8&<=<8%QoC7K!>>ak08@'
    'pKJ<lAfg<z#?i}|nSD$ogh_asfO<6l0%j9zCe$|CSJm@-sZU+v7MhUq!%4_5p#Dqf-{)K5aArvtNe1?$ND{s%=VC$Ygw7<@SSA0g#-b?6Xs({A_SbK('
    '3Qpmg;PS>lheN@{4CaM`R&j1~ATm9w8sZ+VRGV;~?T+(O0hdr!u4}}od#Wb8Ti&yP-C>nZFG?`skJBiVb#`pW6cPP9Y8{~8rh(brhF?b_Su^<~E60i<'
    '^xV=#RiC05$3d<USiH&gEht&S4^iT*r@{hONQGEH^1c~bDLJ*NzN@Fk&a5UlxqBB9_bBa)x*Mp!T<!atROg(2sY~&hjx)a{TC(8)TL@>=U<BCfD6F%('
    '-;M@ok<W&6T$jMbkhe9cpFcn0Kbj|E0@vpu)io3(Q|<YZCu6j?+kSYb4%UJx^n%reYo3F`SAt4MoBIj@{S6}wCE)*rQOu>>`&m$Q>#yM66GF(LApCF?'
    '!^7s!P_o+NFV`6_bDK^{_{3D6e(2Zt*vj13gBl#bBjL;rRA?Dmo7+#UD`HMFEAF_b(<>3%F${X;4jFbhET$;l)*4x4ed<X{u|;A6nxqYzw>j~c@KmrS'
    '7ps}b_A(nc{eHdt>k%2Vo86JZjwNQbTOSX%WazvnfH7hYV(O&9khDT>xBJ!mEgq|<vCJ5&>jX=+)!oT>ZXonLSuV*d9btAOOkvqTOiq(vL1YCQa--%c'
    'Ogp|$>&r9qImplaZi!sE(<c_7@#nC)y$FF!1+;p*LgB)p;7Y#1@$H3-s*^@~re#FZn7Q(9-fdGIXM=3MCZ$hM?%cbG)c|?P7)uV)JCgd?u2;-i&c(U8'
    'd}`oRYRty`+&py+NseIFEn!pWg4`3T;cI^wxsW36LBt{;-3*2hW6skm&s{L1n@q}o&)iW4vO$2zff@l_9~Xz!OK|cDfjDZ4lO+az0OVDdJ5&(JOo8wy'
    '$Z?$?f$28yMc}z}Sr|;x{UscD%#C~4UcpAzuita__;0@shpT-0n$Xd~wAe2<TWo;?`+NV~!07XJ9l85A@R4aYAJ9tZF99SEetR<<Yx4XvnmNJX+27IK'
    '-clSbk8GZXBcq8i_gi=_<+xrhl9jOOVe`#b9NsN@2>C53=J@6-vIm0-r4$!OgEu2w=)e669`j<ecr}7j<&9Bk)dF0fZSp+xeZwocN!B@-c+`g?eP{3|'
    '4vK8|mrF39FIQU@qa@daaqt~pA2H9s<~K+9NG+gk+E62hBrJ)@o>g8)^GPI#k|G%!%DlK?e!=vTxdjs~VLSM)Ho%D|bp_t58P9@>Z{!NpanwB7R3=ei'
    'q(cs1^l6xy?P?B2N(@nz@gz*@g*$9k>*eDdPLM>2@pbL(5NBjJ&EeRhp5W~%9C3AL5ISii@un{E8QsZoi4<MT^fregO(c$2@fd~9*46z`ZD|fjIJ-H>'
    'avCQ|L{z=x6&l!N@HDY?`L;~7cEZ?HUTwLl3V<v$p}JzhiBkZ?Y<su7>uWt%fze8exSeZDL0N)1EC-zU`CxhEwK#;bnI;f+TV0Er)>xzD$^aob#1vlO'
    'Xf!=H<(Ze*TeEIpD}ZN|_lnut?CJ|hP?Dmw>KPF7V&?!voYAww+zEhGl06F{;aRUz<MJ#--gf3Nysj|X(-^!f&LS9H5d3Kbs=t{9DDsdTxH$<x`OE2e'
    'Z_0RmdOX))4G_45%xGPfFO$2CdjR@3s3ic$1BE2D$xQ&kSgS;`S#z-@3vS^ixLur*^35kfkL6*DRYNs4l$`~{MZ(Uy>nO9-Q9zAZ^-}MCni}jXr6HKL'
    'X_a-crstzVHd1QbeEj3XZ=e3{!!KW&>#rZaeQf&QKmFt5*Kb%`i=5eu?_a-t`p1Xn>dRO7O~y1aJ9s=@QcRW85Pam+{+dY|-tY$ugNBZ4W9abbn(cR~'
    '-)lqEG8#|7C!I_C$}kn=pr<;K;jk_b8;#PojEdt?eB91mEHeoOjNZ}|*zC6VRQyL{`^I<~3o{Fj_vJ$d3E81#l8JuL3>}yNqKqced)mh7{9>@1SsKp_'
    'hqeAxRX5}C>F3&9Kt-K?WK&^qh0^U8u(w~K9vt}n?UTWOC->rzrWqTcPNK`tgHW6iOTtDVFh^1IH<9&_l$RO_*Km@jS*%dsj}O$mH-i`724s`!RGX`I'
    '0__VJhq|K5+9UP5Y<9@vFw(}vZZS7RiY{|qhUA)3F+~!;h_w!Z!?wA>=-^u4$$d~=y4(jQ#=(!o{S=!Gmha|n#tz`X1IoTssHA}59qivH-cz_KbdAH{'
    'zL4Ld22v#SRqo4_`=Xs{wOfvEg*1pq2)evsN^vBFn{|J+qAH9Dgc(3s&ov#iRxF+DJu+^=`P~yEbX5e$6<`R1+w%}622;IbKmFDxUN1%)Gm^nH9hf@4'
    'A~vQX)gImjy;<Ch*Tj_Hes93a^6LDsE5ie}ytFxP6H+^f+{p5<z(#N|&#}*S^qIjn9|5^geVwiVomaC)Ku<ninaT_sorz`kbfc+kMpLF6j<%_-C5?`w'
    '(QZ824o7E;lpkqkK_)%x$ur2p?yzADW-{G`^I6b}xoZLE2=C-hw^>MSI4s68j9ZV1{)Pv^CS>Ci?P_UM#Qdjmi7A3>Ew~52Cr!-eYIU>1pc&mPaqxkk'
    'NX|Ccz_(yL?pH8Z&YN4*I&4F4+KbyZo;MBEYTKN5#w(Hbrfzdp3Q|v%y1x|I*e{tQy<hE@Ss_k3ui!5m7Ndi6<MU@$qT5{UD1A#;!3XjA?A08P$@+FE'
    '!He}RU{`GK3Nw6ai@byTjjzL#u}D~!ZrYWNUU8-L1HAP^U&*<_`JmlXC)+x*&|puYsR~OKncbM^!7h^A2@OBY^`vJVy3T0vEDx_5^JGYrByaW3;!M)8'
    'DYiKXqus32T-TUTerA=UF%Qio9K(gA^l+)2Fg+X$y@wNrs+#xH2Si32jSUV(tUizx*{=&13u;hr?@Jdz&UE1mc&rOjk$CY<W`M?;9nQH@uN4<x4@H>p'
    '|1aq>EQw{(sk}ueMr0<KI@CeN*O#v!nzPtYC?03d_urGmNrEv$nD`uwa;I@z!HLmx3vR->hu4~VrC)YirjpfU@=b2u6ZaPEtlK_&j~?68MB|}r?aTlc'
    '=F7i#tr+95YpS}Z_b&B>yz3Xe{M;3>QC5+v+}q;Zrk`s~M2Aj?-2bN8z^n*%;2JmyRQK!GZ{uM(UWZ3tJ2Px<TnBS0N8GGaF}*Vq>B)WKL<I<2*ZLg>'
    'I6PALe|tRyF)CA#gPlf}BC&=e<3fK4W?A!O{P)2SSmas7H+J0bO9l4}ip?nn*N_!mhU6uUl;E)*0tfG81pX%}60~?iQ0|m+kNx_(e}i9OwSJ53s>j1&'
    'iA`n7!d`vWRAK^8D*l?NzY!Gs__YbyJkhPC(DVLw3qM|+d|+liblJ25yM!`Sv$t_(H9pDD7vEo13s-ZqbxaaI7?kC0sQSqCXgonY)(*&rj;U>&ra5!t'
    'k9Yb0uSc{Uz>NrvWZr0`)C^Se)RTcvII@F*$1+kH(ZKTwr!V>BaBQQ@YuB7ncgm{@qujPLKgRIQnQ1Eo0~><LH6yGulGZnisIIq}kWSV+e!<oCYPru4'
    '$gPFl1k#hL>UbhWgYuC$4syw0DfqNp>pu)~osKAF0Hcb+eWGlcgtypPAB_|G(WqEeVmc0^9-~pxz6)1Vr$)UQU(2g9>O=@GC1Nn*IUH<@&{wXzm<;Cd'
    '%5<!Ax;$i@ADbh7)6qdoxG!{ee)O~7CYbVV$OfIed<iqAVNvCfevGW=l7>|iY9l(XDG(!<M(NS)+(#Acpm3(>8Kz%HCDp622vT=^cEP!0vk-6HFE*?v'
    'jNHmd$g9e7LRJvj4pLXvP^%a{kkwtj-@@WmHnNDU)3$(^XSlyI9PIKd_yq?97R${BHgmK;8>S2{cNkdxu(d|!L|1pq%vsHuri`<y#b#DsZ03h^n_5+?'
    'EBIKhpkHQ}pSVPmui?kn2)<myyyz*OGyU-r$)OA?#twBzyq5gBYoqLshtgehNVR{`o{-8e5|C~QsCAj+&!i1<MA*4r9kz4S0nPfdP1TnKhWo4aIY<-I'
    'afMCU4eZqL+*K{urG(D`1qO{#bdYDGHBzCZ(Bc$CXAqsAOKdSN8k%`5sY#`0IhXq|8KT&8ojZY1Z{GpTDI967lCR@iUz7E%@#A1p0>fNYY9l>{RRbhr'
    'RacO7QMaPA5Sx8rYGR#e0V14p7<XPlMSL6|3p*{#1tglS#E{X0zxas8Ox^|xXkxt0+(Lp0-ne`!PbYSv>m94Ipqk^`bDO9~3k1hy6lu3{K$;vO9#m~e'
    '>A=r8(T*R`u`qPHnTi2qcLaMqKjrLVNC=6NjuC1Ok#z>SU!mo(NnoSV5|)P1k;_VuiRz~<YL~6<mV=*b%ZWF7aIf6x!Fg)VKX;161Z%Lw5G)g$@7i!U'
    'k*hk8ZsiJFY=8cY1ymcVn{FG@veHK91%ts^jqP2`VnY!Xx*giCsFCh=i_d}a9Xv>tUN=@f>F}hwA%?eHUrrc0nX4FE$J`&h^d-Cp*q<;|?97wxnzLI9'
    '^(fuGeKmv%OuooamU;7!Yx1bq&3e0d+qib!oDui&5sqGrCv=D^J0%05GKMbq<XAx&b0pY#OOv3vN1co#`rQ7>wNJJiA!t(}czLSPXzRdrc#^Xq4enB('
    '1HV>Vzt&9;KEPw$mqc&YU>qU`3NlPgG6Ed$it-fik;#$o<8CX04QFs9<GKdd=5UXhp;t`@FULM5Y5EFQmq(*LfsS+<ehmi3vaud~H1p)*ll9}Vrc9W&'
    'n;Gl7bz|)8SmVp%P<{0lR@WUh9AKQd$Bz7d(+@`BnsntubQ1J?d%M~l8w>Nn+9$Z6b)#{idU;$*Q;qX5&g&=qe0Dfac~gRnC5=)&Vbmv{P$TDOxu%g1'
    'wx@0ioVNXL``A-6-J$6K1LoccBeaw@O&Cfho=81lgtL1xdNj;&nLOXNsiz$3C)T|km`d98<Uc9>+QGxbCv5GHV?<kC1jn5Qh%bc+p*#=09{vvN<-;*^'
    'nb@{jH@G>d@lrDBo2R=7^<5Ccfu;a%0scwOQSd)i*U`eQwznwn6879HlDXAS21CYa7}3R`N+!Gh;3Fv}B%*!=IznE<%3o8Z&dsskUtwd)4xax)w<8e4'
    'acxE6=gz+f#0)`j6BE1DbmXz+v<k_ruc)W!(CCn|#dwxpbWROW;OTS7KYG^9oEBqW!;y>c6e%^Ru6vDaY6BeCT44eVoUB<sV7W!2M`S8k9!#ZJ^<X5{'
    '%#Y%HoN0-_U9QRM7RJK^Ln#y<9b!c%{=H{j7POD3340;K@Lz*rF%+tD#f;lKG&F9rygaT{xE=5v;eUV)u)RBAgaX+Jg$X5rVb@V^K0v`Xvi86%L&FM<'
    '_y01Znxxy;uE}p#IVT&IAi7u(@sDrNfB^erb@#Yj+;18CzocfD4CC4%j4ABrCXi*G3)@J)nhLnD_eUZpx6R@TAH)wJ;8v^@Kt$`O--`O&M89Tn2YY7g'
    '#u%k<Ho6%#DK9>^C{!gV)sCP_Kn0-m^L*~eEs{LA3ek)ZgBOltOZz6g(z{Ib$?x}mi<u!XMZxOGjYaUnz029$3<f)#$~m!)4EH2wlkcxMdKu?-zf(RA'
    'U;s9XaDnkTLv5S%^%9{on*Ylk`aq-6K4%t`VS9EgwUr%E_BWRm;kSIn5aaOKz@U|C9Lrt&(g;FRQQ&0@!uPF*3l_@+h3Ct1q<`FxnF<y)nyqO<wPOs?'
    '__Q3IpE8;jc1Zd1gFQ@#q->W;^T@T@?kEw<-Fvk#M058a|4-EI2tXJD;Wp;>F;3;*Uvil)+yqgkk1z_OqK$fPZH%9UX4VkIT3bNM`xr`Y7HPHSOm|QA'
    'wR?r_w1^nn;j$05lctZvZ#E`CNK+;Uyi{mV0!hBSCIOrpGaBbaF5H=l;SBrL!0mxechJ3tjWqpX(~o$c2@i>F+sxwv$7&8T-`s$I^5%HAg5{e-hf2H6'
    'dOVctV|b{V4BcQ#`0m8e6v#*2)s#UV48>>|ZK{*e*bSIj8*E^;+`AR{MCi}o^_cMAFEMF|%vK%msCqJ9%46B<dlsM!M)q){jpS4;S<;e&Xco6yIGQkB'
    '`v5Diz1i18;WYBx1gSbwG^<h|i|zFq34qZhlQAeA(TaG4?-C7rPETQ*DfZ7G8=%9qtqjC)3F=|^>_>M=&v?;aVMdGEe;1gRK1T!m$+zR76kc`S`FdFX'
    'eIAdkzP~s4>(TI>q=|v^HQ*#t5$pAGxlYso4D@8$?@gAfmhr*0vD~??Z{ccts8+H(fF3@8Npd$tkixzuN%yA3u%bPizg3S*L~cigU%z72Fx(Yl)SwRk'
    'q7zX1DlZ2eA$j3i_sVgJEXm@5AJSBlhJRd+R`_vhvw_!$FLeBSOjyB`*5&O*vXjBdT60+QYPosDFa|Rwn6hCZ4;4lVd0~X?8t%6ws=i&|q=v5tPsU!c'
    '`3}>jq4J-@0K^4p9&bBtgcyJvFwOTMF@5u0MjNh|qLsxBMR|FERsV?Q<Px*Y+iF1PN_x{yQi@<HBKbthELK?v9LwHJNoB?u>(zC=e!N4&`T81}kZ>Ej'
    'S)zB19e3bjw!b#m13W_WdBr8Jyg9VvyE!O7f(lt=KqpeWIjH;N{r&13x5@j}6f>9+CKq{fS6^@^FcFPh>>uGNt2rh^lf{J`s3Sh5t9m?@FWjv#7GJFd'
    'D%tSTK96LKgOoOQ@}YPB1d~dVxoMmUIqgK(?syzaUX&$R?w}iA*hFQ7&Yd0ovac@hZjP{8-ed6<GWF75puyK~E?FhO=v>}f@I-sSaC~eI?;iY8s?nSx'
    '5XY2!a@v3MldSHlN6G63_Nd=Ky>6T`V;Eaz4p~;SFq~Q=CRIVQH4^GAO69rYgX{Uq2-{Va5e#BYlzzU>TI)}}mwF<P-c3ZpH`5iyTnc&Q9D>jjQX7eI'
    'i(O!H0sVyL!6zQA07uEI5v(ngN*$9o%)(}y?*tYrZp0vcYOK61^waOaB%szS@b7*ej6!`WS4L60#qtJT57s(?CA8Z<!1Vw8TRjNnW?fF6`lMXAI@}ZK'
    '>>Yk-@=ADgsvgf3uwc^v_dmWh-RIBEgBpW&x`5KwpT+PlDgKT(w#T8b!BbhVE<nG(*)|W_m4Xqz>=-)60v%65D?Nj)p2H5fZ|C_1)=}(gTVc^(9&US_'
    'm*2ruhaXm7EQqNeqULb&fww(m873Z+=J@-kUq1XIV;nZ?M4R|_s0CE>qb!M`xY@$)2?tb<6k*$9vARt1fMvO4^uYmjgExq(h0RUbu&N=c*`J1}>5xIT'
    '47-p}N>e9$56@tV>m80aN9-Y@$`0I09~#&$!C*MlBawSMh*`=d>YMG&ab1oAaFM}5>FMe-<CBw4o46-qrR>4|&2dE&v!UG_r7hNt2S*o;CF(20+-$wS'
    'BcxKfU4WRpi4A8Dr_&0yhlj>G6kes>YdFeVHecvLIx7!yg`^zFz4HJ|aB~01g3`(`tz;ORE&T$G?BxUOESqHlakbUO7#+_;#-Ppe_PFEp{T?|PFm1D4'
    '!wbb+6z6}@Nw6cB$X}kHz>S_@_U5B6A5ZINZ+_;!yF@g}!_D*ey#OEheLr*Cd4)5IeIDg<x!*s|fhbP*5=fYNg%!p%=D!f)IL6TBMz~w(BnBPs>1?05'
    '8z<fEII-)4>yp(&?_t3&mxp|ecG~^=751;rg%MJ=QVB^tgH)Ccpivnd4j!U`;bID!Pm}=oTDNy!E(i8cIpa|+6su57I{?F&Vww45aaPJ^G>x`zC~%xt'
    '{&2X~lDAfdA|@EfSiWB%D=jbl%l-BzaIRq&hK&oHD;hLHF>QI%G0dTGmCo&PgVjM5)cARHsOBLDHf8sN*9)gOPP}er69cDS`ZkC&VeVO0@4<%wr*BOz'
    '^oEXFu#U+-Z5lLFCU<0#(akHv&#sKLI<Jm1JAyVbbr6P}bt6+^v_s=M@2M*rUW`;5Wem%&m75GWi17Rd9qJ8Q@$*ekyQs=;YdhI_mPf2p-yvJ*1F4;k'
    'r%g1b%sLfhAQC6#m*HkUQB0M2G%khHaNME4BgxUzLk)JTl}q~BOR6><uhYS7E%0^Csq3B|@2V|2s<LG3s1H_eKHxv5!#!K66g8sbi9UOxBS=TXW_5V`'
    'n9sM-_^3^ZI(5&DKXpU`h80E)LuOXxRq)v4n53-+Lqo6`;=nPI%kp{fWNHp+vTqeIHw1%Kq2#u>Qi9&>63Me<j#mX&8PIxkRcOg76ULyK$OS=7`CJDS'
    '@y(5<a-j^Vy9%-!K?+rhb(*P^dU(G3y!!ZQkrUy*Uo3DQxaoIPR#;8I*@zR6!*a92p6Z+JkCCD+x8o*Ax$*UR6rB=M829M>4|mM!a07>%4mUGC3rrI3'
    'rCnIlF7YhpyMA>?`(F(ihmgi(huGSV@@D8FE=jhkT7Ixq#hHn3yil~YU9RKa$}8vBmb;Ca6hv-nvSng_M)l$u4V%1%aMXy+Sb2jV!Ep6|vU17>ns=i6'
    '<0Zvw^XZ`Uk-p)~<MD?}2WHt)8R1YM!u0!FU75Ut5DJp*(ktfJ&uY2G%iEdPYfUavw22`we;izJ{Ld3#1{SA$*AQYgG`&rOhnkY%<I(xMTFZ3DqtSQA'
    '`x8UGD#sVc*iRTuWj7rIbd|r=4gqj>?J)qx3F_g<kW$fnjJ1;wzlWK8NJjtFalJ@~lAZbTfq6j^lzIWwEK00gCo0X)1|M&W)fQs)?069_3z7AO$LoEu'
    '-M(EVmZrt+YKLL#mKQq%3=un%qE;b4mss0aH-MCP@L9Wd{82iSx9EVP1ocsp-8G4#$d+%&;^)!L88!vH8lfik0{+xy@|Jh~dW#3t<@&hHeo-hY$Rn`O'
    'mBrc?gDZnj6s<{7G{hMXkL6pn_T;@Jvt|NdRo39aS8L4B6nz8B2aJGa4{NS{a^v@F(aVMrUeQggS-FfVNb7r${U({Ribsy1(t}{4){vt(iF8ap*?j23'
    's?L@eHsz4#9|>;f*WpNCsCtlgSKj7kdw00sAJYc#U?wjjSYalC-H+;+(q{YTUj-&gGQUV_fXmxh4xgt(E~V^bJwCzCmevr03{h1p>UzE-Q{mM2+N>V9'
    '_GO)Ja|C3JDhV1lnX#A)d93E`3XV)0{TN0oX@pXoXg=jZj}1KvNfTqh2(f08bcqa>yU6GUVYA_asZQRK9=|0Wgv-6<K7t_z^`pVO>=~-K2{@U)HaLC='
    'XbX9}=K_SvG(Pksf44c};#@bVw}EynQeH<j+;x03nQRxS`pop{miy)9{^S%5ostABoZ5rzQ?<s}Scf;nuPAhCgral46j96x6rMfHLk*+{Krtp#20mnJ'
    '=k2*;7MmcdXkZu*n~XrJxHcysu~stQ@oP8?ulhH1a>S{3q)8PoXVz4bE*%A-6$PQ;mxE8Gksz+v0@SJlD+61`SMT5Bq1nxz1tc(4d{SQymUk;K8|hP@'
    '@a33Wy$rQ8LubdKI%qX320^~d?V4Dwm|%sX8ZgNl95cj-Je|_x*1P7OuB)d0v-a9%vR=Hfn@nfkIT(U>p>GY28lo6gU0R1Kw2IOD!+dbtVCzv^E*@qt'
    '&hz(gzrJovJ6r`ufwz^Tf06rlPwX@_T&ck(V~~K2PYaXJd^F6~;3dFLiQ5R(Qa?2?zs3U-Cz#2>lHtJm*TlUaWag6#mVQmv(&V*xVJ<@_S$FpIYccNE'
    'hG;1V!%aPxR6IQ#rt$aRF>*NvV_VfqZiR`<7$)sae1KyloCKjW*CLEYGA@kAl?{p<eKvGS@fsa+>}6lC?w(Zpw_lsHoD+Q2;PpmOrga6%o`h<i?B9MJ'
    '4YfIg6;K4@^w8^mvs$luqO}_wwWVIh!V{?<*UOZV(TYm65AFroSs5Bpn;T1*Ws;FCQ}tux1FW!JCMT~V2PcEfRE6yi0f2l?4@R!Zrm{F%7K54^)1Vde'
    'NbzWlwoB3o)gmw=*$F?eO;Hw<-&>>V!TkyqFiQ+V^FCoPLFq?0Np=pa?ONkzaGKF4N@%YMtsA4+wrl`u2CZ0n35M?~797ewcE-s(AtPvaT(jV|01y|}'
    ';lAIY$@gUHvWEl`t)xr2JXS!ijxbH`wlBeOW!cc86AUht`5}U8vkWLcN~o;XTEMJCK(?&%cey>Fu!0~DRsks(busN@RE;WBK5r<?{Mfvs_9;<6EkiAi'
    'vr4ukR;QchgpDti2}%b#ddnW-QszrS>uVAuU2m^esEk`e?1ws5_hfK$@#)}t)m-+QR*f$bT)BvQ=|!p<kC|)e@inAH60YH2dJQS#8eV{|U<Ow}YDy`O'
    'a&Goj!9ND?X1Hr-7~!v0T>Uu@!%sesP*Qt*PxA=r+2$aO${}K8W_#{sB}wzc0&Ak`&+g%kUH1D{?MBej|A_Rf!*tjHbk|cL8kPj|@2z&!V7=?b1bR}<'
    '?yvnAJhS%&Br3)#mG?*3^O0W)Zk#Ogx69*O_{8~INr~fi|9&+A!w+GZZ<s3tn4A>*UQVDiyTxt-lT{Uh_O`Gk4<5FT`VP*71Pq(z?y^-AXjk^kXkG+e'
    '#gqkDxe0J?n8AB?vma@3+%5XI%OPgD+r<oMh~>+7KLV`->g9_sW?n*$sl&(j^RPK;H}`7c`F|QgT9xZ-W5<Q<C+B}|0N4&Dzxvy+uXy%X$jI<%yC5a='
    'L3bxQWtPH=snnByGNncvR;H%tR9kprVwvXDt9^-eN9XJQ{LL}HIvJiwYOsv2V;<+a30Pw}IT5Xv?j_~_H*N2>BuA253I2);WmW{>Gyu1&yFYqHdL{yh'
    'TgN2}3E>Nfh7^lL6?%X|OLE@$0eqm$kIcUEk@aJ1s%mCxZf+hCnBAomt5DUU$L(%vdO6xi3^oXdTW30oy4`^$ic2hnr9!AMve_$>UR{6yoT7mrM@rp|'
    '0<VKuxF+$s&NYzlzy0!;46=#zPBjJPLfRy(_<P2D0>@P@7bsurhS<moa>9Nm*{V+Nt46eOGZRB32ei2}73z=on*p?p>K=)ft4jX*`DGQzITM$OXeHXF'
    'WekAVY7&y{R9YNne^&2A?e3xsLDw-N(zeaGTH8q8HAX4m@Z`|%03q9;S(>~!T?s-yHFcsY0Em0VChh35s&R*(hT-@~1<m_gH?NEBY1iKm9!=KejQdU^'
    '%abBn9Mt2DOrKMUGycYXY6)jDSz0m#<+&D+OAt2CGYj_f=YE=~C=@WfF$_<Y$Iv=4y|XtqsHi>5zEi1%n8bC>Zz!r2J#Q~cT0vY-FLNvrL()9mYys=h'
    'y*2m8b6AZ6eSx`ZbjTm33QOY}hf7jm-89zXg_ClJ#P^qbc*$&r#E0im$&Kz&&4}|MGNXmkPX|O)=Mg2rc*JIem@&QA@`d|#wLAx>BG*>el9|XS8XdPg'
    'sNOQycXjKYFCCQLnTL)~8GyDE=-2&j_vof&e|&x0KSq|o17f!diP%fnEmHf%>;w1Th^g?Wq>Q<HNe@u!aaelBYtc})l>%}*_fUW=AP(OQpiRf1?PX9|'
    'M^IVMfNB-jXh#lROw>q=9TMl0^hKAQKoLjBj4-9p!UPYf5e~H}z-Y5S9HvNcO9-w3Yv;OocRD}}1W{{9<vXgSy4wsHs`3PYc&gx)d^}<0!Oif1>Duq4'
    'Nsa+Xv8lCZzopRrIkT}_69d$kLH3}Z83EFVq>f(|9UhG9VF7;C<7wC<cMh;+fL`CFhAkBCtc<*TO^eo%R6m5WBX`SNPWDu;$2%kk8flFcB;&AsKR^ev'
    '(XOEf&{fd6o=4bqhHb(P1<gK#y>R#M{RybHrcG~5gjYFy+_4Vc{0Mx~`~*L~KB3qo@HetuqUmDo8Q`pi9OAeFS?!YG=ml#3_!y)b5%g2c!E<WY4f2C5'
    '7qNVB8ek_|OQhUMCCNH9YqN=l%R;cXD}!ki$LSe8vuqjhsnV$qNFYLps3NpJdBe8_GV#6Amw@QDO0x1ji1gWwD!3GOJ=OmWSoDYOW<eq4SZh3sY9=U8'
    '8j0~LRLFy)7r9edo8lI%tiW4-${-qw5TMpHsanhZAX5*oly{YpQ=mY)(7GklTx<*Y77Oh$O%qb_1!aGoA7bF2dzY^nYhwC!O_crU?4xl;3lr7h^NS^!'
    'FbQ@`cR>ot=8=FIi+C7?V>ZU*#v#=<LQI+!yqb{Q(p?-PitgBLe<EvIK2+A|SoQgu5l7ee3{)GE7)3`RM@i>+yb}H`*T)<pRNWOFMGfWdCm!aM5G<cN'
    'Uz~{cm=c#ZY6vy~m2C>uu=zX7kmBjOJps9d9x1#e-g(9;1DqoDcqtyYFLty&rJnczNxltE8|=xPo3?y#WWFvB7;B9{85w)&D1?CJ$vN0|A~^@;%>Ke&'
    'ZB9^pb?5W|_dLW39wR6rq*st|c;$tIY%FV~hz>2IN6p{``>u);s_=bQhw|)AaE-<hrILlTV^(NvBgdOInon_$B2D~ugQ^G$Bnaq9uAAyU7o>1jd~%Nb'
    '?(UXeLh&Cb^m8WDHay;<O`9BI@9^u`&I(A<;E)(WSP^;l1Tij`*3tH`e6a;)-?7DnAQ~$LL$v8||J1LS>9xtvVugu!P%tQpqSCoI-C)c_av^MQaECLg'
    'SQKVsaEsbi6!sA>Ub5&)SF#gm%#H8P7VQJ>Cr}*q1`F3?b?X&aaW<k7_6e$03MFz3$wP=1kU0$Lw9c2tqsSaZb&S%TK~a?y5k=)Fng<l$19K?a?Lv>C'
    'a}2}diu%|bhN)!ZF-(qOYsQdgFl;3|k6|;6M)BNun+fm_0ErK{H?(Icig7<Fu6vdU!K^2QsOn^~cRD3t8MW!Z{!`ZYu&G53klY`gIR{_~ARAlv@Q*F%'
    'C{291!IJ8cDRp>v95#n#K2X3p;-sJFcwNVMBEd`5@w1R!2B|&g070q*a=l#ax9%LfBZ@qFtMnDTp+9{7y#nf+9NvLV;9J4)`ZBzz&*JAJqFdQAzPSRw'
    '_4o?PIZVHhWc!9M@Vcbvqzhk=^@T5}$rp68N}4OaATNADPrhK1b<<w)1%2TQCi_AogGCCe+L&m}caQOnQ`yD{wrb?~JOskAc#?IZOBjGR(f(aNQU2>c'
    '@vIgwCMj=Oz?E?e6O;7yxYA8!#Le5=#G0X^!AOQ3PzS>;iarECo++YipA;;<-C_;m8D1A-k|E%9)J_Ks7b+ZrNAOK&myB|aguk$zT{Fhjns4h)g;Vu*'
    '@P}z|;^oJwg*eC&Qi#U_>%1V_)4m5MeOf{y#dszR5mqCuo#2p4((xe4IyQ{>_Lm2@B$rH&eRFd6Ij%QT{nM9!{_w5(%ctt=r*D7#{^{HAU!JFUP*sJ)'
    'j3|RCWXd~bZ0{bOo7=!779)jq4Jq=Fj9%^DC8Sw^s6|6j?jEr%BIY^wwa>X{B#n9T-qnP`h9-AxKrD!-SAlxyNP8RstOmXkS@Z5_si4Ao*xZ7@2`Vpx'
    '%9Sx<fZq)!G4a?0ERHbn<>lcFU}mF{W02uJpb}HS<dW;^3xF|q%;P@2y(d(ep(@GqNyR#b{ouxUgqG%yg@$60-iq?cbIjzf>xi-UK>m)?Lx0FdaMJa5'
    '3aGmTu;E6!{x6Wd`eJLJ8Xxj)YBm$7wy{dII)olbMU}i4NKz7s(M4OW*P!p#)%v(UyiCO-6+VPEiMIK1Dtw@00khoJy=k9`<YEVOcdmph#p%?wNh?-_'
    'bL!EP+b`A&a1ho=ru_ZovkK{Y@nn5otut5^+|9~6XwSi-%O=8NMZOtWO3XGCnLfwDpDIy{_5Q)3KZAza$$*T|jquqTT88u}taOBS&0C_cr79kxed?Y}'
    '`jSR~w5G7t&5u+Ya4mu|0r$|>Qrlt;0$c{2`2VH`K^TenDJaAap2f&I<n5vuDYAGCym}Cy>28JtmKt*B=VNu;BN+&={`~s6l1%ETFygQlAr@HGlnLLj'
    '7TecbIs#~)LBu^U2w_C;$B!?cDv>-oJxmDpQodl|IT+!d(8pv9yH?k_?wW8DNKY%AW;?oz0H}g4_zAw{m~G9N@=DA}AV>}x&m&2JcV*Cvo`G+kyp`)s'
    '_%DbU3=Ch;Kk$XzIZ*h7h6j3#O~&zgRs%Kk^U6DuDAxl*LQ|6#aP@_g5*Z_-^NE6kZ$plnG;F}p5YEgW6TwiKL_?Di0TS!(lhakh@)mVWB69<w#KeI{'
    '@kae|xsTpBkg1;i4GP45t3H1D{Ke@q&&Q}T7OL%__Q~k@BO8m?VcfQ-_yMSU?%kRoJ%Uqp8#ko&Ne=5B#BDD2PL**B|K7FMjnfv3vIgZJX@uQOZRt(h'
    'HaTcpM`tF*)cv2Jwj)#x@2sF+QmCn0Q~N#mHluop5OHuRLZ%1W*@M%-;7925taSzr6h7e|Rob)q0>$TYlh8jswH1WJr;^5}hWJTt&~LjY&UN=d_}{>q'
    '`ii)ZV;iw%BeR2QD9>lxHrQ#phGOV;xF_!RzJv`>qxb7!zeMMyJ9?dBiKzlQ4XM?pgEE~lhT7e0T@t){2&7m(oB)lG6guk$Tqg7AahiyOOO0_jj(m#J'
    'W+>5}2QXp;vmT%TM4GA?;FPi+E^1U7s3?UHHQP8<7p^OxoYQ0VFfbukPsHu<@|n}i+<P|-Q2``iPYL;^P`oS&QF-03AvdC=mH|oSZxzapz##uCu1+Im'
    'F<*;_u-ysbbF~`I)$$;+=GCN|B1jS-P9QP^=|s&(yl%@a@&5%VFQ$oOh~sYT_@r<+T@hyzcm_~v1m&j8);f@sVf`7LGhh?(5rB8YX6f!{3$e9J=3dsF'
    'r&b50hoWBtdFb0E1l^=JhWoI)GvarRKodF#WIJ8Fj`cwL_x)Xe?8iS=KK&evxg_Nysdz}J0knkz8RN=5*+;F4rWByUrV`#j8M&l#H4n=zI4;IQP2_%q'
    '|C|3f(Nx13(kf=t2pFAY44GjZA$HLh6Hn^=NYikV0^qpS+s0x+Sax@e^Ll_MEu6Z@wf(}47n`mLK#ORul2i^ll*M`NHrU>+mhH{9degf{F_=J>>zzY-'
    '{{qMt%^1n3Xq=J6TjP7+3)qIi#A5EEHM<vw48VlC$J$p?feKDHZUJ~M{6|p1-JV>-=ssehQAi52!^1j5;%o`&S<R_#y~LtW4+CVJJ-&W_6m6UJ!s=ip'
    'hJw)e{yqKiAH!nHgv{_=V0qrvwzKgVHo$&vOr-w=&pRGdkTXIxkM_gxe%M2T-6&KMC9R3(ZAo!KhXJH8iXxH5;NB={oTjdI;|ch|y^c=umD}-)A=H$('
    'TdwYj5|e@k#?p;6aW$v5z6q^7hA&f!AS9OE)eUATt)LkCVT{>X(krn=IFs`&g6k6fYIcNSJdKENS3BRXVY7fz?ECG(FRuGL_q3U5nuFzyryE^_g+lWX'
    'cu82chy}rlwG2g$U6*|gz0Dm^?K%be5H5fo)25;1u2Tl8XP|VCE;zz+cVx?0Ho-!Y!Z&@u3jTKnbQ;R@94xX^2HvPLnAblU&aDqlnIJG}!ZQS{I7kBE'
    'L6S@N^>+I<fk7LKhN3ZJs4kX>l6nS;L1?W|)V2?l*KO0}kdk*CzKAbFQgmMS<_d6%PYzrum*L=MOrM6($}WeEPZ|fSu7iDlS|yNeJ%&`C2$vxBRgi%b'
    'K0y^ZYCQf#gF5nM)R0QVpfU%w$+X)mr~NQ$r_LdpWZJK5XS{z(6R#AX+@>xs&1~zH*Qb4z4=IPp`N;Y!k<TnsrD7_dq`3AgG5uQI&cRy4u<8e3$7|IM'
    'g36I2xw)j=(<_m)C7Ywhl6R8KmRF)qmaJ>B&9FInU6Sqc>IE@cxemF_kwwz#<_8wUcmWf5cG5U4+A?^1J$T^4Di3rc64dUh?Un(;CeHSTu45VVTHE^a'
    '8Afc*;~m$vCPi%<DDP-^>Llf|z^fkY&6q{RfzjLH?XVsmomR#emZ+6{2@n$qN%zBsf?D^kVVfA@{VvgiNz#X<prkLGwqBxgLXr|J+nJ2@5mQ)H-cDsq'
    '{m$kx9YrgljB3&rLfY&GQf^6Dn0GP}#wexEa`%Xy;EXW$umL`1ts~>PG&&2Kd9!_3-a(~OgHIOlkv;<nR#Dn6crle{Ik}bzH9~ZCJX1SLzq)7+U|Fzk'
    'J5@c}UFwFChAHx@2dD3zEa`NwAF9prVIX>L$i?4tB4hvy!At<wmBD&cwa%j1+B&7zb}`A?bA@4E9-ysye3D_Z(E$8&8cDaoJH5?8RLvmbG6)#+h8#xt'
    'f+7e*<Yy1)r}naarJE|WpwC6aA^}l5C+fdD5e~)x5N->?rYiT`i3X1-h&FSg#qWacHwGC~e4c+eviSbQ=F3}n;RD;nu<>m{^YBY_JwTmNM9e=CV>LzV'
    '9}QQ2BFs@GyZj?+{1Kqo4HeY<%LHEao6`V~Fep;!wk6u)-dT>pLb)8s=m_N5lo<$5V&M(q<2*M6GYp2QQ=PD!t2R?7WQRs8o>YOHIf20pY64U2@9X~X'
    'cIl3c?HY_5_4nUm4y8NbLN03EXjDBZdw54?js0fm6NbCH)c^!KL>Uz^R0FhWF$|%BEReJ|K=6D+rUGP2!av;`w!VMK1pT%#OjhVqLsxjVhQfvSt9~(5'
    'K04L?5iX@D-9n6}2+3ShdBdKY@)M>>B&rOC9el4EQL$4y$cHDY)Aj3uJl$j4!{Wq_GIx3evRCD1s=J~XMP(9UD_FR_R1hPTj8*O3CdNyGLMvjTU!aEr'
    'Vi_3&Y8xL?8eaBI(@J^CEWaBdx;E2x2ubKR)@h(Hw7~R6iwE|K7gt4b%ruLX`pfqydrrdEikq${O(#d|;cT02c)tLIPaeVVV$mS}e#4P5z<I@FRGpk5'
    'IL^q88>zuc=JS@&+(1u-)5fg?x5LP&z(4OI-5R8cVcg}gSuFdw^P;<QxxqCV59P+xp_m^ip=8K_AAoIkfW@?85U(7&!ysjtaa*v%RM4QmD&dn^(i1l+'
    'LOd*juS@>=`#19Rb=5OS?A@cp^9-6D+maH<89HVH;^hm$Z8f512yJ*CD@lg4IRrF=0zqKd^|D$Y4<j_-G;Cs?jA|!c-Je$5w7;n6@Fz_D>#r~0Rnp<&'
    '<pOxG;-%17*~JoWRLl~Cf{uez8gEXANHf&59p(|#(GKNPHWJ0`X4q5oteese1YFG$hy=7?A_X*g39ZB<5`bjKK<W}ukxqn=WT2r0Zme$HB(N*?M_?CD'
    '_!CB4ivc4P6@!BEbs2*X%MySMN=rjk=nDY;n5INTH!m+btXQ7`H6~`G*FQ2{0au9^s+JSnE}b}gwuj?UZYh;5zG-bkE~(q);m$n<tL@zqL-p3DH^i0M'
    'Kj1-fln~@(+ZQseS4(J4^=~2?+B<TzoBp_VFQ218kmiLbfmqET_6&&hAjOu%dZ`YiIRm0Rh$?{u;syX|3m|vy`grYo(};3!f|dao-8_c$7?aMlDr_n^'
    'ce=(G^mz@>opAIVz{_sPp#>-StF>DIua3hzSXOuvt#6!O0tN2Cp7hH7=Rd#NAt}{6m48fG{#f__?qlKRc&9?bR`M~^O4Za+WAE@0gda=JRGEB7x(oM+'
    '$nKB#OZPv}=f-}>s&VHmMCP9OTQYipx|V#AQ|)+MKnP@vsN04O0me*C6|#P(ASRVdR04x##~*9TADa|70wzf7j{Wem(|#N5dk75cU8k4B<}d`e!n5tl'
    'J#~-$bG%8MSVb4h7|to$D!|pYaA9rVV3;7@QILP`^uNZv;NIcur*h)e-4eX6467O_s2DSQ6y<e6C_|5hR#E#+lU7GHQi6QU>2dNYt=O9Za^J`~7vVCc'
    'EMUZJ97Pe9<2K-f1RqTOezqAQdls|<F@cEHK4!#Y2oi`x3+^C(XLIC+SH+fdaIKO}9X|_i79d&<(lcK*tj;FFap*zAq#9E2?Qhkyo6`zHi6lF+0#LMb'
    'E7_`uwC>;c5KMo5^kdmXMP}A<E1&>4B>%jgapS!7%__C3W#INAFc2<Wh1xM{cM%o0#bwZt%H=>seGzozayW-7El<m|ty!o^;v<+@FIMqB>FO)-!hFqP'
    'YF_!BxDJ!1YQiWR5Do8%PF{zI(=^w@phEzjx(<(KXa=z1Ct%Z_Iq`_~k2al#f63>!2?kB4u(g-hVleg34O|nVD_bj!z6f`BT0mLeuncpsPy+=RE$rpp'
    '=N(9MaBT>IRZy4vifwn8ozGS@+_K9D4!>AuS{8i9+K8^39Kq>|y>J}2A@ETv@@<xoaHU=X;v-7RtCz*a^F-898&N80AO*gV=5)SgfdCv^W-g!giNEq('
    'v-q5s_PV8U8C1j^N7vWsE`1$wmrk8rE$^{L<-1B@yj{ApI!M*?ad24H(xipSk;QI<{Hu`Xx>#;K^UqVtA{}7J%tA2w0t{EOO>uOs-6?J&85!ll@1bls'
    'Fhsa_UG`h6Vx>rlI?nFoo*3@8<KO98CZz~Oih2AAx##lVt2X`~HJ<$Uo|<7S1{UAI=|D_TNz=2RgEdO1miY!b39Mh8z`MEy&HHXOydwwLQ>B0XJqL(B'
    'N~IE0C*=ZBF=fzw5kw&?Rb>*>5!79-m1k<erwN7FABO!7gQ5FS(A1rPhu26~2Z6WC?Sk?}Ug4t6;X013rV<zX?RtmB<(7M=(1(A0za2jSQP}8sl9G*W'
    'vm73V<nySj?ww}eug0Q+%*6QqBy!UG)e`GNa9!548~!mxZ^IW*THO6Te;NU>Bpw~I54Xf~eSq2%U;);(sn)9s3zZ5EuW8+w0=Ig{+Z0W!H>dp#m>DAR'
    '3P;j;5gVc>v{qD!W-P)#cW-F7?Az$m@SGJ?rqlJ=y}b{MEgi`RxBU*t$Nl6q@sXBmkvKD0<Xl7Xo}DsZ`W=h8Rqzy(K%hQG(IF5S{-vAY&xG^MLCj)^'
    'sDDs)?i3?l3FlyB(jirZq4WD3MhGQ^ridcOp<ltEo*fiz7^O@qx`I+iXIo+<W0ZT7n6BrW5-7+TU-)<mLtfbPD2-(lQ$z`l(es@O=L8ag6>+>0cdj+2'
    'PFmBP<7}O8i4jp#?M!*1!H$IG-_N#1Cu_GYPZO%%J(Qh2*CJadi>ADj{6%~Qq>a&NzHFld!TJb5iWWF@114>6?i>f32Q<w)s_QPG;Wn?%q4ArX{5j{)'
    '{DwaV7h}#GoVWmo8~<$MVw{-|;5vs!M}S8wfyMALMTKzAf#FGTwrT8xB4*ojX#Pw%1J~6_Pt655e>NmbZi*+8&e{ujboHKv<0n1cE$85-hs6|+pYv2&'
    '<@_3cvptW-PkCKm+7WbsyuZ-jHZl!&c5;}B!U0kSg0Pj?G%>k!Fxp~rpDQVIlqB0+s3J<X(F+)occ#!qjOl$|fKj|qg(<>h+q!^Zc<+d&h>_jI1q{n8'
    'QnW>k=}laKX?P({H6>yACO+GkmKWeuS3;D0k4-_M?E*kV5oEHn3Me9Jikw+0Z_Y2bHbmjO)!&x8Q@Dl;IEI&;EHAla7o9`uYCg_ndDZ3D@C3w&hPC3x'
    '&TGZ@+tuyteeUi{A%jw&t!dbnL;RV*MA9!_BaJc!dmJ#0-d6@dO709za0po|uRs$orFr`^L~__V#|i=e+O~<Al&~;pv4?uF4CY%ldjAoU<-uxmq?igh'
    ')5lunXdwp2+OCrmw0?bqi>U;cA*pD^v8v91LXpxdV04Jl(Y`Eu-o0_Po4J1K9M<y6n9L=3v=^tx*-^`vnGOmIfT<}M5bMn?GMi!9a1Q~6tfV-~<#6Bo'
    'w5GY<2Gkr)Cf(He=nc@a*H{m-*f_4K6yP*Jhdk|Y+PG|`&9#Z^a?57i!@|JNfOH`lyir7Sy4!-GeH9XJ*ZsStcTiS)_pl<u)F;_avA{_fB0-1=jT<z-'
    'w<{hfNDxgSZSCn~;iymz-80ERO*hfBlMPaeHmsA_IqKD2{|F^$PjC<>Y>^O^0)EVRfg98%A{7FWsqA`SpMrNn4S`ZNmu^_v%u0h~lS{TNY3IKSqHQkP'
    'v1oI#$w0?kJriw}aA0^KAsDE=0!UTQq%7e`#_6uXd1)MyE+O%35M9XyBn$i=1soG&-)dQymy|jLx!s-;UdI3t63&1@{33)bLPX=grF??35UN@|l-Fob'
    'K9nGx^*oNA0kN?^t+)tcvi`IX&|zYg9z{M!R690Uz9q!_y4Ci~-6qL|wD2a3JmiWj8c17o1LYQ=1d(gR-2mj8A3s+4hf$iBg*H=Z(SjgS_veDVRkWQ7'
    'N{NC~xtKEU<eI6RQgbbpNlV+QAQA{w$9;cToiMoiDOG$rHv-z|7y)Btu@sUomNt${%)L*DWupoW(dXx|LR4ZICgyHAnC2WzjFy|?w5ChW2}4Y&a9(2)'
    'lqYZj_6OJsiJ<C$Ny4xs+W&|34zuN_l>AusH62v<cu>Lkg%SQDrS(0m`^DRF-3XwZ!9rKO2q;JfiDkBn<Mw{L=~vZjWKDYe5GiRu(}vQ{IS@bU4aPmp'
    '0tFYA$GORYIOyI4I><5*Jd6TL-Y>~daDxumSrpF7^icge$Kv;|L0>_R1s03a7&SQ%zg`V$p0hyXF>3NLQX&~_dk!etZ~-PCB&GOnq*RSLn0%0$$sj3{'
    'Oj~;n%g4ERpoa`ccf5ay-zUTI@okgQPT!A+WCXXU(PxoH!!2b#+=e5`Ok>U>CF3n+KHi4CMW$)a;-rHlWj07oLFE<M%JcK@{YE(dha884VEmKiOves='
    '?AHW`RaF1LnA>gq?1!zk$#A+l!7=CeerS%{`+0M^Wb?OYx5ubC&7?VOv@&gXc6d@CDKj8xRLBVH$R&jgXOX#@gx{iKR7o)z7p?gjO<6iYuSpdezldf1'
    'gnsY0F9V*eZ*FqHpgZcC;!4I1@=!^OrU=0crM=^NjDY(fk>6)~!NdqrkS7BB#cK#TIh@=T2_?(=1ynZxTMyuvItQ?^0aR`LIBbBmF>So5q5gG##W?*v'
    'p2PIGLIs8tEIw4S8_uW@;RPRz3Yt@M?x8Xt@-wZ5cSsh#+l|y8F;Y|sEO@v9thodhT&@!siYugSULh{QI2D2zze-3=oDV@Z>vKqPxleE`zj=uy<1CK1'
    'ZH{#rfbk|;@Xu4}+ORyJFy?-|JlrgwV7jflg>GJ%`eTOW?W0hj!;&>?3w*N6oj(WS_IqYu8mNm2@fGs>OJGsk<XBcRmP)aRX(>f(bF}<iuqvqTz}6}+'
    'u7Yq06W4TU_^k>zgVwTcFChjG!Z|>t7_hrMpNIY|T5u!*n|6H(&8J44feZJS)@`dMhfvnN<TzUFZgQwbFqAHwXWX&$SUc<%8ITS1Kklq<iZc+77{ZAa'
    'd$25>t!>Na8g3xCWRNMbdIo<yK}O+ox0iQ^p3I7y7{jgJ{`|apX6y^W?EW6MVaIVJ(qQggpvVe!Qm1vLv>Cl~JRrM{8=)@)G9RBNo0;ieh-622RCXaZ'
    '!dyn}AK=-!!3R;fne0X;9*qlto)a-?qnWgkeaTdI`Vu5i8O}7Xt&{$?m$8n^J&;qJMlDo%3tN$Ny1R^eLpA8jEo5DjjF=a&Rv0cd*Ij;YHd0(My@8xP'
    '>2BR+-BmF>r}f2rIKvHq-7^^SX}3pAu8mpK2I8<qO+lm2E0a(sdE8h?pbdq=W^p5_%>>kDpwXAEtxf8vQ}vx>Pt%sRd3J>d39s!|D`n<E$U@Ptlh{=R'
    '(7}1Bo4P5<-|m5I&iCqEQ-HURo1!m4%{C}jECxFCdTy|1EI1w_QW#2lX%!*P8I(Luh&<C-3jKKAXRJVkAwym%s51Ke+VOJ*qzJy^c6p>F^K!JcnZx9B'
    '*0c1Ogl91z#pIlJeql2ZW-QobP3L&An5ye7mH*uO0%Hen^g{9#?*yY|k)_iHq~OreTGZD8JsU62m(ME&WW!_JXw_|Xr)RsI(<T|vDl(ujcF;oc!xM``'
    'nYPumD*r?m_(Px=)D<UgbWEjpLNhufTPxbQsvDw&bL;ngPrm#y9keHd208+yKzJv-qWnzNi}mvyH=MqJ+x;MJNLI8NFqG5g5@(94SW=Mo-Ey_AcKa><'
    'vp*fTKnt?RKz$#oE$bxUo3L~wDpU>XSA=y~ETN*~2qbV2Y=JFeT|J}kY+!tEk(bHIsdn1jhTUWh+GMhzbl2qYM|_{_Q6qt4P{xRoj`FAcqPZP@8XiY4'
    'o+A+N=jqx4a`fvYprrMVGoaW6wg%5;iy_WS=nZJID|oS-j_%2_8g7S&<+0K)^}qc5atv#0#hg#ryEP2g{b{J620aS{p?r26qByh4H&obUdpaDlq8^tp'
    'qJ^>Zv?y6kgiG=v5vMg|PvJJiqb@*5Gd`LOlow=ABAuErrcT%oYc`5XOq<Zv!M@e%nqH5t3&Z<G@?~PX-*11~EVJQxs@!G$^S2xlon`LalxN^hw@WN2'
    '&$S8g@SR`<uM9NLT!EK1t_T>?#Q@OG11?TGV7NV8(MAzANvkQ4><Y+B%@kqJ*p}GUW-yPKkV8aXl%n%gSSXzeb#4Q_1TECnutw#R0L!926M#gC(qb46'
    '1JeJpq{=SN7OKOGXwxDJnw#FO?cTjLR}fAa^iv7Gxm+|+2W&kN9sUdz)`Q8X0MH7yj7fU5`I=}shI<9yF7Xgskxk3p+>pP%uXJtU*8TXgVhSgZE|~(i'
    'C6B-r=wv!r!&L+oK%urx0o`^^>5tFY7H!2;EAh{#kZq34*Fc;8;nxpefBF3R%Zq>f{NcBcUtaw7>F2LsUVQ!X&tKB30O*&cj%P3x0=!2eZ~VKPhufuV'
    '(LelMb+<ZgseBCOu#JqI0#D6IwFi|su$CFTC3i?+SKV*7x7FeG_9rG}G<pos`f?sbPofLVL}QQ&qst&*j$t5#kG*@E*uZwnKmG$EB)wzZ^Obwj@x{w+'
    '62XNCR>xXdn#`|<VUHO_M@Z3e&-l&#us?ukr6II^WPnA8CM1_31kxtlE#WzEH>{TL-G*Q9z-%<^c5bgh@caH2b}Jw{0U`1xdA1FY7p>}hD7sq!nTU&!'
    'wcy?fxNR|UY;~n@7!F+%-B&iC6A&Ohf5{nE0+@G&wUJGw!G-1t@PT3D<TMRn7up=b2i#ZNn>T~IsEvCYRrl*LP@^N?w`*I=LmP3x{sl?{M1(a^GiHJi'
    'c*h(<%Wm5OKn)1(7{NFNN=8eKmUwSGEH^;)aR-k+V4(mLFkZF$?eVyLl4E%p<3)jXUU<XVn?Q+0LZien5{vX-(%WGB4SvL5>yOnfJQqPT^a#r42+?Nc'
    'Yi@D1=VW*9-9duU@kJ1=b<%n?+29L$Ik=7G6xHFK6RIXcP23Sm_;M^a1G3DXmuq?Tp&o)EcYP}M_>>DZyn_{55MfZjen0-0$i-rzRVSvyx8PZX1x$kQ'
    'M6Q*IoZ;ngzH3Ib+)5d}G*~4L6EoqE!*G~PE)t<QImLz2^P$W>$3&Y@Gb$NLU&7=*j!rITdFJq8HYfL_Gsze@X@QVwwAw(EJssU0ldiibOKUEJX@VXF'
    'Fm?vUYYmU#@}Gf<J`s1t)H6^#1EVptkoX!i+dGB#3|*I0^Ax81+2#Z$??6C^_ojF$>!;fNt@`!%ub`G>piQh)q%{}S!=fnM6iPLr0vBkpdcp#28q<!S'
    'j}hH5=V^~bM|y}GU+}xW?tkjh^D`i`F<R<@LNIq(5Q>U<nk_N2-k+9xr?EgPS-gG%MKT<H+tDeavA3dHqAxGU4jG2f=k|OHp{F;C)C@-uT{3K40Rjhq'
    'igTCYL^cfJ?Thm*Nmj;myo?U7E*LLQC%OWrLl2b{jKZdkHns&`%7$kx#=fStuxPF(O985`0--FH=*yecsZeGRa_!rZ?!a}+05|o}KZ_o`N58e66VYM6'
    '+i%}xO&44#$<2j`R{~g*Znl^(wOc~$_A$FnK+;6^1kilE-gtx%wzvHY_#4ZPL?QQD@FFsPH=o?lz!3l+rBW=9iB#K2Px3=jmLo~XZsd=VXj!SMooLWI'
    'KfGk-D9FT?D(;WgwS@+Dzj5lFSN6Be_7&8Aj8H_?9uq*4WO9dyX(WO8wpLpClf7E(`}Y(MetPJ0312`BCFE2hm<z^YIw&v(gIK3x2UPttq2K~Wp>`_d'
    'DO55k8^fcuk^I6~5Y07I3fd2$UXRmA-MTzhLb^8xgsuvp!y4n!lTlznVcalC*8m30!)qWU+x~ROB|HZMO6a0Qg;l8BF=);ATv{7$KUU7|R}cNe5)$tT'
    '=<c+^?mV(_21}Br4_|;Xkxbk{p)I656zL;@3IwO)VaK307B&Gog?3EskkBIBomMMPK;kc_F;2tfC%T3z%!RkBM?ky64l6$T24_?X9w&Hj!M&`m!oe#$'
    'hp~~Kn-T#o!)&J@LEUvKyS)wsv*~h}sP1>om0f~pkqm_kMRbj?RecF19Kt!0NZMnrLP8@$4inv_u0ic}4SIiC<tWil+EMo96^$8}OAA0TbJqch>lRMw'
    '#A$S}002ock(#?*v|(D4z-qhoZ0v^!sB0oAUo4<9GDJQ<&%jN{Qgz^DG0VVN0ZKs;(;b~f(6(00uZ%DK;oTAJdY<w7#@&C9RQ@y<q)khL?TjE2xVyuH'
    'DyZfNv_%M^nIIhZ+nv)kR^_kLz93mUFYCXZei};ZYQyCay4KsXY~>$@2VWS^wSU<8fZmG4G8|9IgK=CURJ#ZI<6ESE0UPK7EH{{34KD){%23Gs1=<&6'
    '5(2mB_OQHj?e`q)(a?Gb4e9DJgslVF4>W{phK=ix-R%e9+$7s9u;qyEC6t?f>G|;qC_@&J``2?YU4@KKS7e|BQ)2)Z+4xzXkV(z9*MiF$n%$($mUfH`'
    'Sz%#@8#IFfH`EbGt5STYtR4JI5)GlK^C}7D7YNwljDS^ZgR0%N9j9x7PCLCkoffuuTm9rN`cQiy5Cb7PfMw`SRKy$Vfj5_>3i9+jKmo3(DsQXYKL{y-'
    '*4EUK^B&A#vSbFsy1CN0AB^lyAnm#u6$lDX*@2+}LktgtpC?Ko98EHaCMK&^HRnWc`kO~IvkpKTw;7&<JGFKY8S!fOdh0aho9*`a`uPFE1oMw{N&DI#'
    '_(=Z)A5rlhbsy=@edN|Hx8ZhqOUrG$0{iXUhw_iK(da=OeSQhVX<GC+|6~KX?N?M%FrP3LGek7e^kt^T2}&lf75CN1vje(4pm}=Qx5M-Mm+P0LJ&O8;'
    '$Uf2XGn%{#?z+Gi8|A0Ez4+?h?HBi!gKY^^Q(zapUsuXKI9_Dy+nwYgtxZvm*;zCW77lni$75Spi&JtIFJX}hfGP)+Nq4(TfFs{b3Y%}QE|xmdG@$2U'
    'nYASV+8mHg_LOgb77*zjQ@HfN)UI(cK+MexF2uxd^cvqlgRh(M5*$bqztM}|K79S~+dqGJZbn_V4xG{`+^afc$frN;k5dBpZ!aq;)T>2r>N*GOLNJn+'
    'Id^{qOiAhiOO~-LKJ<;uGVK^=T;9zAsR9t(HPOz{5qVwY#Y7+1ZNDBu<fO@b=}sP8>0xo0IteA*ZTD`$f?eotzjSlBTtEoUQ|9(1gcua!vk}fBfV(zF'
    'kdda^)KXtSIBf60R3ps<Cb-=@a132QNjNkqCpGsYBVr()21I!kVsP)4zyyP-tH3A=s)!{SOJ9ZMGogxLiowiv;}&_Qi+Gyx>{WO^!mEsC7_GSqZMTFt'
    '_ac^MtoACb=<qM0HH_9>IiUx#ukOo;91;!M*c>GCY0m9PWc(MrBgqL9s18<3ce29cYkM3*5c$VkI&v2pdNfR=pTO2Lj;tif=vg+zJU@{nL#De{l{<=3'
    'c`!P;>pr|u;A$CDXhM__ZdW;td`00-85WjPVXizj2cQ!kQZ)s@-57<_^2Y%nhHutV76A_Z{MW5z;d&|@3A|<j+L{5x6d>?ImE;?EyEQ6Ko=K~L^tTWg'
    '`pONRQ(zVt+_ye0e7OWPA{>R&FE>y`D1-DNCpx4fufMo!y88I#x9V@7K7XnHwfgw=`~S>@gouINWH=wU!yVTTOY$%rw+l>9$y(86fGUQjm>OR=0TArD'
    'i{);p?m<N!!SpFz3&H3R_XaR3fgy{2Y=@IRid?qkeLX7I61)LtWR2zvx=rWh7N!7#_8>&mb7Y;tu_4h93bvZ_^Lq=-F)Lr?z^PyZdN0?*Eoft#WT^6Q'
    'w4tJm5~?tFtj$rin-J&%m?sRPoP}<KvJKFryzfxW*({*|<T2Q4gEw;I1|kH>1K`U5?)o1PSIIocu{ihZcG=%RH(q7^nZoI0sG2cc9Dc&UgJ%|~t9f}h'
    'Vs3q87?bf8R>$p6MC50dhnp|Rnsq1Y<IVQ5N+!+iKc#bIuSIEKj)DVz`<D-&fBE+DL$&|%&mZz1`-D#TF}|sWZ~x=_r_Y~$|56>Z<0WBxG^7R3^!zZ|'
    '7ix?mLyZ(Tr0~)8@WFlY>zBz8NiJhzhyeY}5P`RZqK#%2G(^8fG?msX;>AQh!gW{{5lT*0fsBc;4o%lTgr=rw(OV2?vWOPf=&J?QB<7`+kFe&OMQB?@'
    '8-RgpH5{ov+<sUg?REnCx^DCYSsBIcY|?f<3{Da&0k3aP<mQ}T0;-Fud1p;`VeR)$liq+8z4wP$zAjY=w1m9W3kU76xf}LtG!MkdrCThU5?y=C5-d}y'
    'hWh<0U8U1IR7WlwXu@@OIm9ZB8j(XrZURFItmm!|I0OUaD<5!_tX`l|N=;|@fe^M2J;h^U`1zN=e*EQ2{<ARUwrhl#UNJZO)0^mgiPz$*`>lIn0R8`f'
    'wZQT@(=tZZ8{#}U(>?k`(5C`6_b-E1e1Vf%cm|qx#BRBVI7S~4ny(I<f5T}0MtXYR4`72?oly*Q7vm$0RHedc_uIE8*zfmSs#a9JA1O;?IKRc|1jmcy'
    '{>#67uRi?c`{&=MXDYK(x{sR+*A6^7(W2xd<YtH5SjniPJ26o}oK|y(u_^&1ySFT*4by_e&3E(lyVG7bpMHjo_v)`7e);U~rQbg}HTK(=U;puI>CUi1'
    'N3}&>I$HeFB`pLaXfUwX{V$Kz0oQ?BW^ViV7H40QzI=!dkEh+2&%b>9<wNxv8#fsw>r1ZDiGtZ?t~dW$-F*4-*J}9f<A;C#H0e$gW6Y$XIvM#|SphrX'
    '&FTKk3Ta}sQ2^0bfv716H^in5fk5*G#^}_Od##};O_FNT*|_gdkeNjK?sC7}ol~V7UGvA=`f3=N!$i}aA!Xn?nC$jy+E^aVP>3F0$o@D?6kCYuu`Py#'
    'mzPizi@0`L**Zf?l=KOxI0M?ntB5iXzyW%jo^+*aF{%UL*xZM5YdavXnEg^4v!kq{$mPkWg66a8ks6z7{Q|aA6gNlmx)rcH;GQ9DRV#|^yBRk4Y<zuM'
    'ZxeG}F6JXe6tG08ve-Cofy8`l2z(8-;DT5l;6j3af8zJ#%=adN8zky&PuN3I7R@3xP_EHIPSExmfLUaPk#+IDeNi30z#)`W$|iAOe~2w^h|_0<`I$TP'
    'XgU1|0=-l3CPy*kP_ZR7xfBR0f<V=m6>`4Pqrm<17~)|LIhKDd1D5UPn&rFNAChXsIg+jo&t;d4E38Z&TJIjpgqto13knbRQG}FEGd~oqr(6O#VOS0-'
    'h;Q?lCX#4smaw@5RkZJKCsIC&8l}2iYJa*}P9%IT8%i`2iOCj%Y(ExgrvisQ7ZBVg@@S+2llUHzkx8rcWW4fF9+L5zD9Y5i%yd_~Q?+G*Ny8Rk>UGCc'
    'E@XHG$Yi{T1En;Ros}%n<`Q!=NhN4WN+pf+Vx_6<TFWAJ{-M&$bi68MHx<a|UIoTo6i~T9dBjvG?FvF|vi677!}3raj*qJ$m(V=>ZL;?9pnTkvWem@L'
    'o6O6b)9}75V0rf2WLmtUUli$hPP$}H_QX3~kPtPmmM%Fj@I?kEGWf`I8O!6<r5kL2LLWi#I}MMRN_w9ZWuH5nl(g}fmOfu>e-p~V=f7=KJccFz7Ily8'
    'rz+mig4keulkWE0_d|IOReD?6@|9QI`|e4{vo9xmzi>^0`<R=Tp2(|VyX%im4|(>>Zy!H@|JP5|-`pSm>BI93AO!c}CuP1DrpJv6<gh4Ua)H_DF(%n_'
    'c5>hphr=Y-jheKrG#n<mZq#IF7Y}JS!Msp)vQW31TV!)8-=WGRH;z#Sv6&-L+2kTgL-?cM>=@eQIx(UkSsq54+$ctt#HPo+vdOJtbU~<auG-{^Ft#K!'
    'J?@oFZU|Ev{_?EZ<d!f^NvyCuEZ-Duo!k`n!}9RDd?a-uxx!VxC1<CqlX1+6;LeM^gJ<YsSv+fWUGuR-G$`Bs28=;58%B23$6qt4z{%=9MI7qu9!ZU{'
    'K5l;;-c@h5i?>78LTy;Vri8<Z7(^~^j_$4Yrr$#d;j3sa$e=<t7eL`Yoj@&mu<%Uens|5_;z@FSJnon8S#Z7G-(uduL%()cT8_|ap3qoA4r*4oal;Ds'
    'i&u~?f2eeffP7KieE#Lz@1MQ`7m2RE=ZJ<+5S>iiKl~k;rk9(4_`6&rx#_e?MbH&E9NY~#6>7r*sXIKh$lP~;5zl_wiMW>N1Q^~Ohs^<k>A-paqPpDz'
    'C7*jUY~E+Grsi#I@GhcZy?X~$4IZ}D(mgpA$H^gOYTlJbbtAbjM~^tn(<dKf=*018l3hkArM<g3ODB#A;W{vv=}z-@Q1HxR53`q$QJSwntK%A=)RNwF'
    'Tj5fGP20cUZrm!s#~C;uiF9+kEZTfkm{7eHzhUOTf!QqwR8mv)9L;i1tv%afehR`>JRC4K9IvCS$$5)4>$q*KeVKd%tQ8N)Urw6hfpH(hx7_CW{IUox'
    'c-u8%3hs~W0}t<`9}}7c2!Y1tm%szXq66p3LI<DmeTX=;ug~vhW4ng2)C4Pb`ARq+B~5T*0;l33^Uz~-&XuStUNK$QMAW<G2D5Yp1Xw3I7MIYpU^H7s'
    '<MXk&{!6H=26vslgj!spCDdrg67+~WiyF4Y1<*i=2#b>oZmpM<fYaxX|5&1_Vw0bgf&38HbXOL`Y0cz1R#NFD?)$Laud%IQEj+lRlkAbeglQm@Y<iA$'
    'th=^dY{2HU2Kzi(mclyQ{MUchxiBlSU0q)g7P&C1vRy4M2+Ogs4FzaWSY8lTxv)sStS$)aT$pvNuGSZXO)ji>LD%Mju+4?}%IRu*LAc3<`7Y_|=7MmW'
    '3-i^{)$Il0E*EZ+osqAwz^*sH$ZXy67U}BC@*<b-k`*p4%jZi?BzIaTFU#kvO(fS=r!LFq3r@(S{ra+ezUBm*klM~%md}@+V1rWI*~{|z$`fo{YP;sL'
    'e7^Vu+nU<0y)2)vKhd%lm6}6`<;F=aH`O!H@@|&*+oRWX-t>peU;)AW9z@n{LPWDT4SOiZ7j06==eIo|>05@O={AGf=ah>d9xR|0-Skjtd=xpdI=y!<'
    'w{@K@$R-4)fz1j32+=`zJ-%NCcUQd`g0}PbsPX=NF5d~3H%0mT9=JMY<ePZ*I(b`^e;794E13akc}Gx%;T44$1(t-Me&%AygAdRAOw8-qxl~+!zG&*?'
    'W1xzzaw%SPq|;_nL_E4h2A@Oq`BQCk0Dff>f!j<U31KYkDT}=u3UJ(mX%+G#y|(XOL!_C}JB)}RF?n)%TIYC@EvcrPcwRyTd@hw3>1;ifnpo-DI<eAO'
    'F_9=)McX<t(OET-DH@mBI<e7NJr$Zh^kRx6ikOMW)Rx@ViHXkIg2Z7Wkyz(!Gm$75hgy+%glsz%DVd>K!N$Bc^@ycG7mwxP7ReJEom}fb{*BEhfBYM~'
    'w}V5j&mY;Io`yD)RwbLu4kJIeL%uWw8^MgUB}ocEgGz{*sI<v>)tx=>$dCN@IY2ZG)Do3}`2(*EpxCcVG-UHFfG7S$TuK%{f8Z9yqlv=%wa=WlMR}e0'
    'D4L=?8~{7e<2KvH@wA`q2s@JLC7PIzW{w`smqmw1B8l9cHjm2+&%`$;;%}ZYYPMP?HmfE>==Y25a=*+EP+!Xi(;v{CdEc*M`W`$#x39=JI3Uefb=%kO'
    'dafqBMMuTArV_gJZI5=-PtFV-m0c$wqX2EpqdUqDyB;}(bGfMGIysfeK`uNOC0Q*G`3}(?&kJe<BWbYM?Xez$8!51B!uRtLGa_k#EkJ~b+lip!2}W07'
    'Do8DczyxSOE@^mTyE>CZf$DyF-}@l#Wb_ux2~@nUYbrlG4^_GwZ1glbag&K7g%oBEpZ1HT>&kr|V9{D3Sw0^;j8><CxY<J98Zvmz4T{w~5BMI>=V9e^'
    'oLq)Y!New<%lPb>T#Wsm6p)*Xy&i6s<TTy)he@BKDOr#QA(g%f%p5Rp34w^ylMe%gEfr^>YZ-B}ioNB~!MXp)$e47I0W~6Nnw<lBbQ1w)vs_pvt!N6P'
    'vU4hakj`b5VcC2y;5ft}qQi2%T;cN}ABILt2FX`I`H&^0jzmOa)+fhF;S<qoEr`L|A876%ufTr`jJdf;bY2TGpU@|~!#~6;<N7A9m+D1M@q;*llNfPq'
    'j&p<M%|nQspD4te;IL7YvVKQ6^P}6c*^{(vf1<ljoQS>SRpr*^Z~(fRT$T-oWCily^!hN205W>!-|T9hQOU<k|FA@zo)SiL8NpL3d7OOM@Z<|@eI%}~'
    '5XQY*9{>2a2{s!YYso9@0g4RMH!wH(^{`*w-Q}{H=TFrWSvu!tL@f`!;`3(ldV4oHn>vm#sL4s!ALy~x?~gzhNQ-PN6^l!wl#Nn{?dq1C$7>K>?2*U;'
    '_c8ip@;j}I^-F-npF0^JZ^z{x;yaOvVhgb_nYiTnR=mcxZ;x9nqTxR8Ryvq<-J%&>n{s*fNzotu++%6?BM&p3MCI387ZQ8`MDzK)A0FvOdJOv%F9_F)'
    'OkS{&wAP8*1t;$R%3B=EqriYAt95v{T=ZZc2|S%>YT52RxsMyXHZ#|Cl3)3?NzPDrmi_VX@hU6NzhqMvG5e3ywUKDvUCCA|qFC}t7VGVsC6Mm@@xO8b'
    'wh$1XlV2m-J)C3i=A1-Nc9LxZ6hP&Nbk(niN%(6g*=|MlS6a+&1l@gnS{(D+r;{9ERVU6}H(@tJ^=5EuClgSDztHe-y&fQKaB|3WDp50<{ALU(2)Rtu'
    'F_iWMsr#Eg4KCOWpj?8ERgUHDG~ebF<^SO8z1%Msi7zRa;%mH3{CPpD=qnO+?$+rvU&d|IPCp(WEmuG#H>hlr1+!W%{`mKN;57RqDbcsQ4GLFmXOb({'
    's<~YdTJ3jcK7NK>Z-&ZJ5C!VZ9vCJdk3*QnuT*=y`u^K5fBEp2Pv5G4{POwRhq)8UG#oNQr6h7lS<gWhOedX*LEa9pDdNpHabmpcEZ?}M>q>BHws~Xg'
    'kpUX6B-NR0fq~84z2Ev+%hxHy?vIP@@i7;V-06na?^xXDYn=m#Jm7}5%?JP*d`{ZdiFBJwGiAmsLe~Z+l~&TbFk`b{EmyI6V#F9_xeyFzMUu~R0>ecZ'
    'X;87s&^biTj_a&Gj~Ltw=ip)q8V6_2!%?Ju84E~L>!h=#k_w1LX1y~orit7>wGROv!;lpu0_5@yQMqGjVs9!XGAyFO0=}!YyHHN<aX5GytmnurRay!b'
    'xWi~EDpNzf>zIGA-*K)9FuA@H@|_lAS<_Kq0m{lc%Uw8bZ~JQN4&tR#NveqI7Qp@%egk}OWnAmSw-J|!?Z$12OS~-KP-95ms;PW<@Xm=F?v^Qh-sgJG'
    'Nt0?ojc*E#2X2uWSs_@~6l7@=%?WEBH7g6k{#=@hTX&FXMaFKY;?;x%`(lZKMmIp4d%L}_jweV)0fJxm58$_S|D&??&sgIoC7B__NJe3WG(U*=Oh^KK'
    'Ua3auTD`(en0!LPJ3{a$Jfqj1$SN1xwL2ZKGVPdLf4z8ddso5ays~37-9%cLMkvo14{yNbyQF5DYuwjrv0mLfm=|^G07Q(o4x~ga1yVD{ez<irxOy0l'
    '9wmaP6vU}Dk=;*gB8%~r#g=Z<tad$g3LI)HBo}w>T6mV9Z!gC{O$gePu0xGyV}F3kHvM8?VZ=%#8>svcYN<DMI6AW`KK<&Zub$lwD1FmEAJy~mbVeEy'
    '_H16rhdvOib=E{l^J5>NSC$*@S^Q9U0q~n2#>aOU*u&hPb*pE;(naEYVSap=BEfv9My;FKPnkfV2A{J3e|5;nq(eVGY~hg6otXVXxJcm(?T>$9R4Gau'
    'HT#{wqvO6qF47+yuLE(kxXq@!X7($*!-N%AePuW7F&=#j3%Hp@j-2s;EPuHDM<^KJIcA3qSb0I%&AoLamVZj~(z5D@I)RTI*Nt|r8%>}C0_-1Owy27_'
    'T{Z(R+cts!@%4(g@nU=09QTiv)^m$V)#0Ln&$R!Ciz$4jHFKX49q$d>{U7cPrv%b2H*)SfvgQeIuG<@E2A~h<TF(LtK0X?8Juse5x}F^vDS2QT`6IZ8'
    'o{H7Y;07i+z*X1d#)Dhe<&|&1%MeHatW2tUFSd~Oy}2EihTUvW_x+J8*<yy>7%v370batOS;pJB)-E?mX<P*ahZ&&Y1}21>Wk6rUH)-G$FA`r_A|N@7'
    'N7XzTPft4>nXs%(=sL@IphIHL^$KW!|9ncJ%@jH^%Q|Rt1@yQ?vpWzbh_T$3Lu@V}QqgsWvs1WMGF*EB&ewxyympEg%X&GyON(*QFZyi)*G;Bc*4$fF'
    'Uz%z#0w!>B(mEMAsx0yDoP+BhAUJD-CFZh*0uRy{S!u~tXvg)R7hv(n%^ttw8F~rU9ieUge_;ttWX3f@r`6~XUJG9Gs=2&{`o-xn9SW0<Y}061^L{9l'
    'wD=5IoNllf(4him#^w7S3JGTgbDiYqydmO{So1dsh7xYb(kto$2A?sV#VtI+2!e-9QmIg*&Yi0E!q@Rwo@qjyF#u{V00kpZ8B5gO`3ePoC>+zx0Wc~Y'
    'CjChyTQJV(c-U$m6BR4o+??+29uFa=5vl?J1sg7tmlYILA49c~t<;j`dkR%}It9E|nK2$zePR<*Q^tduJTON6>K2!dU&NIYp1@hdaOLLVP#3Ul5)a+V'
    'wu5p|v3>;JD53B!J=cI&qk5{EQB@!k40TUIP$YtWG)S_M)kqQw$C*gA<5sL0H~)USxyR9kEU`+q=`4e=RNhf!0{e!S*zTA2OPGvzm8>HZYtDp15Pa?%'
    'efInDAZc2Ymdzj-f8HA`@XIZUcTZ@^Army747LH226>j&m@5PEA<MVJrrMu|Tl_D8)q`);fvF;vjitAUsN$J*pyIjfNURU&Mq&?}Mpmczi44Zlp@>cL'
    'Q4$xMgym6+jj}RU=1@NLgwP&O>-A%`J>a6eT^@iY(m_ln*-}^m0mw@L&x8aA0@|KFe0x5!;1Zk`45ub=w@dJVpJ`1f4F!}m-FflLZE0jEiPHMmsuXND'
    'EQZAo_AV64t96K~OA)C;8C4|q&Z&T%19g!D+nOH&8Z{EzSarzrJpO>Etbqb01{Kf!7VE%<H5T036`;9ON6DWvgjhAKaT*&SIAvgu`33<ZSjeQw&m9Q='
    'j(Y`-g?LE%UF$b%CPIpw-xsMtw6HPCo57tEh9UHMh}kWrd~qn6<a#nUZimxuFHNk{R`bL*mxK!|4rWtlv*|S8<jk<eV{~)6fz>>F6h<Be<FsA0q{H#0'
    '?DGH+YH80*Parrzanxx~o5gGQV_T;b-0)jXTNm#uU5mLh1xrBjQ1~&(p`>M~`bwy5R5M5$U$rWvlBbY9OmBX>xw{O)iN~NmrynZ_;?mm>_gJ|l8uRb='
    'z=8_O&^yuI=i;$Im83snwLz*N5$W3(Xxol=9y$afeg^0Z4pV#7rS`~t3QZmCgI(A8inG%Og@{J7xi2xXBWspZOD3m-dmsiKt?i5|E@vM87KZ;N7EAp0'
    '5+?xS(rgYH;--P>ql@PqO(~ytg@Z1nmy>8aD>`>jh6HQ{+;?c@Ui{OSk6%82`2CmPs+(UvfBxlPKK)iLzWi3LK7IVBPp$*Do{UHnON;rmlEpuL|LdpU'
    '6M*%npFg;#kuJ!F8{BP7YYKNJMoovwqSsXp5tL?tXf8sCtn2_{wEoys!5f^eI9Lg*{?_NZ`R)2tmD)l@d_0aki!%Y}v};<;ZZ9I=p&FX__jl1W>(>14'
    'w7Ta^=`Exh*EyDquSN>`zKF$dt|?foOYgwcHDJ+@Rzizm8rI;lyNm|rT(|8Z484dL>1!Qga}A;o#4BONC=g3J_(d%5I4)tufDlW;B^R;KF1}bnoMH*D'
    '<wuFVY94}5yad)s2D@q+$SF{O5_QtgtH$433Z_UR>0x&r(wpK~3CZ1St{VHX3#tSe>kK)_tA;=L%}QWwi7<`1zI`RAP=QAqsy%*rufK*sjr&a*uI1g-'
    'Ts50y1}H#9zX3pXS8w(0VR4!PWTT5|)zvHQe!IFY!6~-AQ+2M3a7olb0Z+@gLM%OMQ#_%Q8Q4SU!KwsegYmIZ#2Rz&G?@M_l$HSA&tbKUsLy0w{rvgM'
    'KYk%4kK{*)vW|6utkt|gLjQ1x47be#KjkfxCf4WB#5V-AC)M!o_40-aXrf)zSALE)+%fK$gfqC{7SoC>Orgxr0W84~?vdc#->tUbNfBV!`}K4645I(t'
    'Xq<cn(B~Bajz7_gkd@X@F8G;Zniwcex0Wx%m7ysSFV(J1#xW_ni$l46a8E{Mp~&ZN-+u*O?>YyH7ivc}xX%N9{I%K;(4<B2F72oy$hAN+2Z~o{M_gOi'
    '0;wD*-k%Og{s559f#T)qfW!{~nH(tInhvP`0U(<L#T&Du>YUdCH5rhsxoecIa3xTi0}1Y$YJLEy%Ybz7@o7aI?F58^IT{|Oo4RUYh37cIx94!W`WhTj'
    '#tDu*z!BHr$TE%*j3ckXQDq!&QAQkf4UR73G=_2XH8`e>(@MrM*WlPPPVgTBj=ct_DdA{sctxAy$w<m00t@{Qe+RT7Rb5rpGjRRwptuuUYR|`Y(Gc4N'
    '%aj&U&#%o`psl>5v<V(K4b|tafRSaGP(2r5#1$~A3=^v70*t%@rnsK8t&>ivD_~3+MkJlmSHRdZj7&OZu7D{oCL1aP1B|@_rY*ySf}Q}=TmjRSVBr4X'
    'EppY-*p_kP;3#6CxQJ734;Nyytt{#ecQL}7x0}<O@|{LdBXAoCNTiZP;Q*{=UR$I*Sf)|&Z|JA4dNzVZkPJ3d&IDt5U4#v8SA5TrM1~PUR{=v@zzAbV'
    'Vq+k|0n9TR%HrhyJyj4ZTc>oa2qW20K=}>Dw?mJq*QenDo)I_?`Pc-)1HdswW#G8qu6LIk6jG99=fjJAHVx&+5vs0%X)hr~sBTXC8^|VG4YyFP+O3L4'
    'idTd-#&1KRAEzSFA#wg)qQ>tMl+O@DogFQdm^_z&bz$!9<+D4cuP58Zn!7}OPodF}#D&#A5)ABa1rl<C5nc_RLVgw`7Yt=(nwr?A=f4i|W@G7)A}+Nh'
    'Zz70!%Qa@$1ej-w{!RZRwQGO<e2mdb9WI1M;}d1yudxi!P<glmH|GXDAx}cM!=X~|R8!?}v4XnQO7aXrO7LTKT&@S2rEY3Q^fUA%O+_$c;8;`~fXzIx'
    '+cI2#1<-PwsF(@z`d@5dG)<0L*x5Em4yqcou(<{~&9CyPP1fG7j>U^?XX|{)6S3;H`iH-JSUv!oIdKE@ZWLyYkFHi>R|-!=k^&lP<TX2|tD~MJyRr@k'
    'X&76KJswtf%b`+mv+Voj=4Sg7?s2#2vKMM6f`(#Siq4x`)tMZvUbhEjZP2K7=H>3q;E$6OWkj+ZBS+b1lYaS|3~$S__SiKX8Ao{y=Rsr{L?n~nUIbBP'
    '5VkRSFW^}aT?SE>LF7dcQwE7OKOM;Bfwg6jSboZO&E<h@${@C3kjn$xmO+|mtn}{lvmhny4BXtIB@Ge7L>{@I80@!q!+~PLKq>p{*H6Fx^8Hsnc%^7t'
    '?H;{0;^1cN@0O4dI8p+}MfITg)XVKlVcmJ2qXkEsATo})mr?{L4*+9PPgg+=vE=gHU;=6WZTH>eNb$mPa+|u8sW}FLuM<$$jDd95#Nm+j+;F(#^4@da'
    '4AuK#1B@naK#YM4?n&g3UCodq(_k4V2k@ns2>`VhKQblpZ8a5-weX-^RGRNwl7Be`;4*TJ)XB`e9`@VeP+<|sH$4z;lv=N&<<Lo6RHye3q$waFL5=}y'
    'Rha7Ar>}o|`Fy%gtciy~L*?v8c#AF7Z9k7;4fkk7I|t(9-!>eg+feZp-}^TwU`|>ts-L!d_W&q#-5pz+MGJY-Q;2Gc5l!k1MYhJ(b43G8P!m;6Vb-P9'
    '9C2u0k0DxNqjo~kr$FuScG%OL!F~8TR{!*6vY-9+^9QI5Q%T6Em(NS6o-Y7VyBSnp5$6gx3Kn%bfSQ5I#_AeK&Eww9!LH#yJu+Mi1mfm-4i(&HuCsa`'
    'YO`^g&J`_dW5j5q(f3oDR6lm5CDpc3Picq%l{KCyH4*Cggr)N!Tg&G}<9*i<;fLQr^CDef+k_!QlS*mGcKuJNwr*FGVawn|h}Cf|kaPPWfh!GL$Ip2O'
    'GTMn_$%&xk6Hv>-<#b6rzW<$g!^4+~6^`vLkSl_mGHHALvMqf~$0PtJRG8*7q9*t_(Lu@W44Xu}l=YK%+1AgCCr!td-5a_yN|WTb$qa(!f!5u)$0$~6'
    'b|8fVA)`t?o)^`;tc*!}ToTbW^6Z4YrztNwuPOV}c6<AX)o3e7LCD498r1YUm=TvGE;I3lZ*pTxWpUpO?vxzj=XK)Y^$4skQe?;+f^UJz5!@ji$D62K'
    'a)e>A-E7woC^x=VuuyK*%6ft*4I{Qcf+#$q$Ptxf#O_BBrAL%GqK+d^p{nzU$n4f{s($U3?gIKvkSd2XmLbKDK?;J@Iiyt#DSr%75~LYJ#<B>qX(@$f'
    '9@5_j1JF){$(Bx&Q!n&~KmPA8NtmjK4$Nc#ZLWv*vo2-z%)GE=Xmpi0RrCC~5}u!PnN5)8{KJL_h~|Na_{AL>4-iD?8>s6XFlHS#R@OfN=v$Z{Cy~D&'
    '0nc|8Ld4M;A7&;$Tq2o2g2^kvLRTka$4ZWs>PoQCO`U_W!b6^;Gs}9|xc`I+hklc2o1w{QTxlq`<jgFP|KhY?V|*x&U7KUdq=oH|VTK{>a!keNgY13`'
    '(~qRdSD@rO+N$#_aNqAvoA<!v<2!Bg)hGG7CQsZoV7|8|TQFS{`8qrgX%4m<b~}HV3Z<IaW2I~B5Tc-o!lH~6>0a^V6H2Ay59!;XEQ{+Nl~uS$vSLgH'
    'N0K>mLAn0?%g>o;vihm`@I-X%UYG!Ier-ebjwb-*CNIbeD#(r59R|$}2dFyt7P72Ro*43?4-Bbkwuk-|8Ao4Xr=G^~&k$iIMD0m{W$R0`1-`lgSSnXZ'
    '#{G^;XYBhINc6e`f6{tYy%W`Gv%KwZs}Em)e_1!Jd;uZQ)o!^Nf2tMX)oeRTh&}Y1`#o6FA@D0JFDu&^TW+a-1r97yO1r-I37uvn1-e@e%Nq(7hg`-~'
    '&ud%A+6>QTr|}aPE^_9qmWSh(x<`}%dv|>8oO+ul2YQt7Om>fCWv)P(@GN?qU;+Z1<_erKC!YriV+tVGbZJ!3lQ2Y`bVy&*p_I(eV??8dl(oL99fv;0'
    'h|Y|f-0qh!#-!FB#555EV~1X!)>!}j==s!+=ny^huuvZQ<%UQA+^JfPHV1*V@%C)c-f(k*y|LI2{o4CdTx;C3cLTXCkC2*WM|sBtBWuAh`WhHny<M)T'
    'bm4Hj%=;*VZtWf}K78cD%E_u)co^1I8&vv>%}SD0%bjEOl(@ELIu|nZXJAx3;3cKrU7Y-^P>F{LZ!s(h1Mk(_*TZ4?FghnNIu&Ymu3`f-7}Ld^XQ~H#'
    '3#@zYkbSo#=e|!OfVlY3#p+3!=@PPa!$`7-<WmSDPzfhSpKEt-UA?T`eM%%^&oi{fhTDX)O)>HnmLGC~$AR+EG5yRiiKbx~H4o#jk1|9^O_k&^nSsd0'
    'C<AG^9T-nB1A^h@GZ6m0s#+;1WR&J9P6}p%dRJr0(3a%!U?o)70ufx7qA6Uc{@R+}rZUYd_E-(KNll?hU5H5MfLg<iYNtrPW7jpRlQHSb=X;E>W|V~j'
    'Qd75#WJ^eIu542y3?mi#7g8~<1CvPyO+h6JYw;Q^!x?hKkX=x(lAwfZAT1~ESv3P0Qh%;P4Eq=m^^fU*8Jz8ersN0DmT9J>l|Bfe|MRIpTuBh#GD8r0'
    'jx;l5>&$`pV*x>!GW0mCdt}1iop!sWQ=p}MQuSEYh=ecH*3$;u@0SOv_b~T)+mh*(KF=(yw}U5B#S$m|W*KG)69(O#KO7cYs3)uHC+%Llbq<>lp^BDa'
    'eW4FE?u@X?y%*XT*C%|MZtyKQsTG`(ksnz#vRsw|p>!8DqLyeNg4-Yf?eSIuY7ckmpD5?i7lEmsi9dn0`A=>A>+fID;4Gj-+~^CIY6jwkNd(rhu@B45'
    'vA-YqTm2rW$l*^C9QFHzRNG_4*IrB6c!brLAf8P1%)MxyoO1Gf4Aay)#6dNUk9}Lspa}Ev*9x%Q^hulJ1l<&H%vl^*9Ea_SLiPe4`^BA>4dvtOkYz5s'
    '8}hmp<`1xExbUCnpjtDiz8*9jv1QbFl>w^y0aRR<WlW<bW18zRF*vpaX<X;I^cRM8lF~_~{-T<7fEVuq%A%($jsOhT{z9x7fV=JX2!Ri|cr2P%Q+)Z1'
    'xI2RPP}VyaZ@E$=F8vvCUkp5x)+SbEB2pH&TQ*ofBM-UqI!w4GgxfQVUvsor_|C7>Su`lqT)C^~NysN*3P`aCp8KnvL4u+e5B(l^sn0>hs(8TmajsD@'
    'F#K#MxiqMtWUOSshr=CgZVAoEYhaLc!zDiidUEHagC3sCb9G~Z)4hXG;f*^qKxkriAfcAaT==B<({jJ*SGPU#)ExTLez`@$<|olUiRMp*VQjhb1sRwb'
    '<7Rm?;4I**0ty`Y;!k>t(#0xZo*1)$;-~aH4wuy+;!DO^yguy)EJEfR16{HO2+19yeBlX1(vmq4uJ1(;ECx9)cd<H4Ppxt9op;h{4yrW39T)qeN7ft^'
    'm7qU|C6Z#K{$<0f4DE%`Ex)5>U6AJ@qx8^9<ZqWfDgbBYxe%!6#?#{~(}btAs($_bTPDqQES01;5YLP+SOuz*TOP*E%g6lV!OgFL$g(MYJZQ>~UP*CS'
    '(>`!XNI|TXNh`1r<zc+U)SpxV#3%-_GZPWdozhS$K5oc4F*-1OW=4Fp7ZF0%EhNVR7$<Zv6l4Q$j{RZg+!2~<o0C6-N8=7m*`r{6cOKKQH14lygSg_O'
    'AhJcq+Vv17nLxQ^{?tCf&a0{$tON2?=~x;h2j(*9lsBHtyw@m)QA6>Jao2_{Z6dB0DZH6HqOG@hVPr^f0m%Xm7e|+RycsZY=lX#&GC@)okYrq$Z2^g^'
    'I7zK9PNyO)TRH<F8IZBe2#fek^8v?0y)fWhtWr>fVzT2-yG8$YI(g7N^CwZj;yO^I)s1#9IZUQdg)z5$P56Yn{4@#mQLAUS%g5)%!$j7x9=NPqA~{>{'
    '1`LwKcwsPArDe@u#KkUW7Q1ZtjHfUkpkmuT_z>It#~s&Jr>bq+6+G9D!vbh*O9j<YP=$bAtYs`pHHH0W_mC2{&!k0CVF}MMs;}j`3^&MrnVmB3t(yG4'
    'iN9Yjq3%1-H?2<?T%C(0g$i4)Nq2&|SwH4K)?9@QU$*pPVv7O|JO4dbX(IOx{ocJR`UR|??B^;e{D=3L|G8^6mLkl5-f(L=9wBf63=jP(|8>i4>9}Oy'
    'O}}@4O<I!F_v?66HrxH}wAq@geG3e>7!HfmZBh*||9x`9qo30Mex~7?>wFVW8sNe~I`eZb@?Cel4#Ep61iMUH6{;TKXJDv~TesAx{+E0GzH%dVz(_y('
    '54p5r4Geev0!hR{=NTYmJ%2qZqvLc?BG?1gVLh!O0y^DP?g*``k6(UsSJ&S@eg5V558ra6_zn{t9`a|9#Caqu8A+T+lIM}QjCqIV(KATuJW_nkIV5!+'
    'NuNjJ66eJFcm~OwMG9pZ5XqcJvgeUR(mQ(|sX31%lioGwk=pY}s%E72JW_WSsb%d&q&0|r71?OM%EUUs=Lxa$K^E9%{`Y+MQ&{tgh}?Vu*ZkL97?RQg'
    '{I6dx!MvOQIC=eg`z2jr?g(;9;X5!I4R4mmNt2QnFMcGKa|im(Y9gm=KGGXJC4m7Mj}Xd^K>FR&f4fD(>a5h=v#*~%e)#(Hhw87teEa>&=MT?w?2hd<'
    'sNvoW8@P$8*QfP1mijv6O$s@Az;=r7K}g(nr}|q=<%K%J?q7gwKj!q-e{xTWwZf#{dra!B{_@E!mA`)f{PWA|!(YFCdCJ=+qC!&Fs1mYf13KWaKMuKs'
    'k}P3TiD+!gCD>S3qly?M*pQKn@Xe-B9l;{e7>9!KW!cYVUGo>(^b0rrAOA5?547a7P5FnByVxKG)lKY^a?kVpfePMk7J+U4x`!}&r;pu33db9)tqeid'
    '?rMD!Q+dUkn-)P}J#@22#D}?njaA?&x6G-YuYnF~g;zEf!~&x3T&&~i!M9Mq^4!sz`vEh}X~AV7GbF5Io#2{OzZv0iK!=y*;=YQ)qX%GN#PoBZX}`VA'
    'lsawds9%r;`vIbYK-ULv&BNec>fp<=&15i*<gI}#>S0N#G$+W*)%8>$J_~E2hazi0H8_QSkBNUPw+sgi(?SbLrZs45d1TPfv;uB6QmNhH=r*5p?@{-Z'
    'oP65G6UjB!a9fA_<qq^4WI$M~hAnt<^4V+z&n6A(7Gphnc=iY1CKNcpc#CjYO<@$T_r{x3sh;?-A8y>u@E#8Oe5W_DuC3?k^MkhA|MBme#d0@fRaU)J'
    'tX5x>I_LX)TtW<4kCnjl(QRYh^_mW1V2wB&_wHsG_7(at!R`9rA&dJkwWm8pRq<YFynMJJocwk_ym1Tpzh4I8w~zAn#U6L*098%g*ygLgi)GVma%eC1'
    'r}xAgw}b(P{QuH%(s9Xjg}11G+qqr+7K+e<!3>VmRX)5O7mFoQT0gvrgj1OxPPeDApsCryMzs0iW?Vp32Dctp;4IBMXaZ{|irJ~yV&89utuIW!8>NVQ'
    '0Ly#TlmJ3Gdv|6J@JK(0)wGP&OtJ8C-=+n;XK?tCBnNVa<5UuNrtii9bR;VHh$LJ7jOaKB(8~(LvT9y751I1PGHJA)m-dC9b7{rLkZ$IsN5P$5R9hBp'
    '&vZG9Uh~oopCYC{GuC-=WD=M0lBmf*?e{Wp8!YnuYb1MW6xFw*eO|d6<$aF?8*#lA_J~aG70hxCyAtiT{scG_CE3NJsNmh?!{p3EtbT)GB#+hp<c{4*'
    'CromMV}2*rT%pVD&1rp;?@`$#$9qH9HYC>I)xA4SN2yCrIhs%?JOn<++j%X^1RhRkS0{ONoE8T!l)wLPfA5aFw_sEna_zt<#j60P(aXN>f9jV@*S8g!'
    'Lx)3Y)2V2!tjX;&FJ0PqD!x;EaQM7<fi%WrumT`_!S*?f^7d5?Z+>(@%>zsy14gyH`5Spb`b|Fix=y~`oFDz7H_~Xy#|P?CJBI-48)ZFCY3#M)!vj^y'
    'lHly9Yt4FFYi6X&<E^6+w5H_Kl4xP1b)xl9h3}%=^m|RRM!UHYDjt!#<@JdQNV%6-Loydlv@Jv_iOw9Yrq=xC#i?;`i$y5$)Zgv5vy!r9NmG`b?qXjp'
    '2c@W+6RE<uny4E(DKR58nMD&dLMJ7QQl-^uSpREinNP)vQ)t;hO}uArHQ!Sb1Vvc;$W3LQ^yH0TESsxkD^qE>t!~^65}cm4D9*|x`Cg;=4wUV{l>3ST'
    'H336SF<7sZtxhp+_P{(w5V4MRQ))+3Cm3V__IE8fEh~5!%O_!{a)upSc(L@Da6@%@Dc{7qefvtQr|wBlAX>pAh+86SKHcE%WeH_#pOl#buxCc%i!I*w'
    '@%`)VF0KW#&;a(nt9rBm7|3e?-E{Ojwt=g<EBLLU5e(#pUo}$@HiauukJcc(7ri+r)I;5b6Th82mu+3MLLr(IiDo9wE5P{RIH8}UEm%Hm%~nN97xDA('
    'utmxp>>5pyfVELQ`}FPiPt~tq{_^Sbm+H4K)!)8<`uyQf)n9=T_a8re{rc&7{sH!jDjAGLlu~@Zq-N9Fk<~hnTd=|;b{W*?x9eq(r7UlU!_tiyl*@cr'
    'VshX1Pa`|*pM3b=2I#QC?xUtd+0@DNU-Ns-E$&l(YPmhg?^wLEZnl3uVTOPc*#l);?j*F0;%jd1d4&ha?zG;$T~5SxqMa9>1-sM!e$q<AH=QZG2b_AI'
    'H?y&A;;Pc@OY&RdMP=J6AG@aEWoYn2D+naUpuquR>;Cu;A4Ii|;TSmrJKw41#5}^B4g#6{(R_=c00r}H@ltSCqu?+<bvek8&gaLXZLVop0VVd=OJJBq'
    '@i?Q<H4Uqy!~#15cUbO+&2kci*mNTHjr@c#16$-oa{C#KVL)UJK06BXjK<+=xnJHc4-3qdeX8t?tGwx=j!{!EMx-iALPzIN1$f|dz<9T7q2~ZcZ2Lu='
    'wia>AXd))Uo^w}413mXZF#7mB3Y3e~<s4J-XyOfcu5AU#Qac4(ZrV0?Ox<q5Iv!cmJWp(2nHtGZ4>=YaUfR}P!iteix`@^VS|&I^Fvj9|I-cAajdxNk'
    '96h;7y1>YUa%w8ag`o?wGWn*S%7*Z=lBnWG2)c4ml%$4nOiD>wbEq&X+{YzaoZ+E3>Txb-I7S$sZzZ{c4%-DLZWOdLg5^`8GYy|uXm}~f7T;T)Y59CY'
    '8(B$o?vU*a`^Y4Ptbp=&sBWD0y}{VQY<bv83}I($Zamt}&tOozdI>}(4bkU8!kQ_;u$PSJ|Iff6Und45m4K2+cVWvw_wbLBY!c)ynvyIrT@-~|!5c;-'
    '78{7UISm+gNN)nrAFF5IKK(w5@_U}c*<^CsXjuoi{c!94|G1xBm2JaIixTonvXcjA*S0*DvYwGEUAA4vR$i*kUXm<$!MaW)pHwq~QwvSkv2iy9u9oB`'
    'r$ooLS&CIH$rZE{4N3Na0Y*~A4NRRpY<Fa>q)d$6_JA>8kS_$TPB`h2A@FY1uR-D3Kd)9dU_~5(1HE7<H9*{<I}FR|tKQuhkT8SzJE}Xlfr8!PHe(=g'
    '$~)!4)3&r$Dw!7DbPw~lk5!u<zfY<)quP{&5nNL4Wk{#m?!&j=fBF3R!?)`8%hz9i`&KQ!eEsl0zx?*`m*2m9tJc3fRd-*%|Nh}~wfOMQA1YxXe>A~_'
    '(c_vftr8+<Qao-SE2pgQAJ4EU5{?QyJQgCbXUN*X_wDsCcUmlm)p@Erj~v;^S}NQV6SN{$M%6}gu{sFI6&Tl2>8Wrm<soPfG2Ld~gLz6bf<h~}CBJ*Z'
    'W(M^i-OH?iIPFTScnFLx<G5ZO$_;2kIvGfhIN;96go+U$Q(&!}mt<<MmOKc4<WS<$V2Q{d<TNxpw~TVa($ON4rm%?I6(|C6wLuTk1K5)?`f&Hd>f^`i'
    'sVxA=Kw6=E%rOEandTrpD;%qVz%C<Idl9l*G2g>YhCAoBo7>1I`6R!;EJ4Lm9+FnYc_`$dn}On56EMj$P`e(L_c6oS9H30vX3qn9GOigUtz*+_jESMQ'
    '2_|^r(_um}{tSc7-KuJffToUoA@pdX>FXa8rsX(+r^!o>Fob0ckzsHi1){u5swa7AsaRUZE{&T&mN@}71c2~NH?X_&4%twQ3E=qM3bwF*HNu%0k(P@C'
    'Nes!r-tg-WpWS3`JoamM?cK)#I*p-}kX%1W*5i49jnydp^6^W1^I!j4)H74Bgysq;Ql0_xYw;|Y41*3~J-WcUxd&Pfcp{6Ew5nNpbZ`c_!;E04Zw7*M'
    'ES8rrLPavbNHU*U?9ySXx{&{?p@2yF+VDPExa|9fVRH{hE$~=RE+H*Kc}0h&CY;M9oC`XGma!~@p(NdGcHFz<blevEre&<zW$9rO=jN<bxL)9kGb3_N'
    'UIjWe>q{w*#(*WycH4x~Cfe$D#0Ad9C$`6t@0sLc0TWA>B--+?hj);45qCX<ly|;1p?sEZs6a&gzIzC`m36x<FG+18VY5-TnQcIVQf@#Km^B+oe*5g`'
    '!737X{+zjybsJbz8>cBWGbr@5I|tJR1sEy<*eML!GS2`&@mB^Yw1b}LXAuGi9x#22Iilm$a!r3r)t#csY=n$6teb?BHo>@Xb{G+8Z2_Vw?-Yic!LYnr'
    '+L7kYl4Ze0#gCJRH9NM?i%VYMN*2zQ0a!6-wz!dzw5_2cG`u^`iba;ThSZ?Kc%20Z={3?eHO^US5!`GS{S|K}6EbQ53>~qB)*x1~DIxg+q-WRTe(4tW'
    'HEjv$0Em`w=X-!l3(<zV6vW$dR|u*s-o#<0LTGV1&j>X(HklHcN&|3kkB0^N88EM4dEE5**34krK*ntwT_;bY3<(K2E{;nZAKVUXW{OL|WiZ<@*l5w0'
    'AZ9Q^)eOKWy|=Gyj1^5BeS`5Fqf7SY^!~j&+wWf=ABTQ_>9VyVlE^odmPvQk1f?#>hH!k8jaLKsG2-EV9ga-ekQfs(;Y(9)rH*Tv0z68)dNMo(oWO(+'
    'I3oq%8tGV9$hNJl=ZqS+j}Gj<j1Yo`Wm8jijMCCQE{|9)tR&xdEN|zwCjp$@k+utTmW?+T@w%V?eiMt?J1L!z+=owIn8zkoum}0lPC4`KSZN+)SqBRu'
    'ZAm1Qo&b@qxQx$=hr_V(VVwSprR`jj50S63GgQ~Au#4a-nvo9lOtmEfD8GO4@@lJ+L}UaPCRj+oc7PVE;wMZT(zb4r1|}^;=^620pM?H#3$v0%B}paY'
    '$WmN^f5UszNKqs_nHryiH66^6_{&^AJey!}bo^vi8m0W{U=juInk^60RXBki3~<U_@K<wDx{G_i=@K)_yhI2Zh40C__`O7*EFy#x4+hP$r~JHB2up<7'
    'mAd%dVCpN1`@jVJC~x$$B0=qUr-b0iWsAWrq^cb@v+}%<Wm6ugtD&srX;S6Il3KpYAdO-W?$6X>UY2focbWMn<i_Okp+Z$E{A{#oQws7^aa${13*|E~'
    'spL|I2qmZ*Z^8DHdr2%9zM)8{Dt{TN%y-NC83Yqo6~zGf4_}btJDSKqORJx((<Si`f(ei6q_<J8STFbe;<VbHX3q51#sTkayz|z>el>efwUIo&rc#F^'
    'wKIyI*EgHC(~Kf`+n)P!Cz3DALiHK`c0pJsA503f1))pYNE?DA5r7FE`J;77h_5b+N8&0wXBykqe0>q+RIv(HX!%?{Cy%AM<MhV$_1iAA+z1w%@yVU<'
    'Lol5GKb*ZulN?8uCiquO>I9&%m~OW8&D_Hy!T@(SFJFQYkUC5xiOJ$LKu`cOB`R-y>_Z<~`Z9;P%(X>&WLjiSeN6tbd5dnQrsm=7DrOR7`uj}Rr*D7o'
    'A~tiaw-)Qu>S2c7zhJgDK~~Wx`scOlMwHQLMiT6qAqZ!1s%=pNr;zYhpF1(a_3}dUj8*;Imm3jCBXx`^GMA%~W%RKb>FgZ2e?}$OjJCUV(=8Sc{~M;s'
    'ax=;dr80qX&ImAqxnR4bmH~9sc8!q!@&BD0k@Bwyc4H6QHXg)tyQlm^BCXUp-o58nM#&)2*HO&mh%mPqN{7v8@(T3@AoQ|mu5K3ewOxM&XW*eYVC9`n'
    'e=}pmZrJ2*)6%UflcUXj+PK2{jV;iS(WEZLsvCcC+erBLF3!}Le?4rkLa0{Wvfebx^OibiSs#ZN&U>ukEvy4jI|F3vx}Je)g}=O<DwBh8b2^9PRW!xP'
    'GdUb@XrD!KR<2r0a>M61_ieY|KL9ip;mn?()+u}lGuO&-1oO;G=9|MJpdLGfo*7K*yMxr43%t3$tfTKg!xLKH8w63rlqWnt4&T%R5rSSeCTwzqvI>Mc'
    '&)=U%G2X_ZjLaOgn|{AOoH-k9P4`)Im||8c=?=G;z5`&GHt{^K2f4`J@J7pVW$6SfyP%DGDUmwz|EA-3Zfm(`jIK;x^-YJ8TN@|RHkvg|du@K`54Ra#'
    'v96TTPBz0XiTL5$evN$$4(KFYcL!u`HW9FpU%nKw{PG*hYrhvjS{#R~;kdewFl*qwwpKA(f>_?$uXZO)CBYJ86ohxvQ7@kIuiuIWK)HMYeBmV38{qd{'
    '3g`s>MzgnT7Vtm(6}A`<9pH=QT4?=7zO%@SmkG43`>Q}a<K9gRY%BtYch_Ttpp6vl*2|9TaHm7F;VMjDT|l1zRKnN>1!xse&jm#~z$73NwBD$sxY3v#'
    '>Kh+a6&Y_f!y3-hp$Fj^)4Q<J5{8Fiv4TCOTOj7c?He@1kB};BiK4>VB-gmwC;r9TuMWfBXTH7j*;)yZOLWCH_aZ8y0<YYWo;gZx;?_?On>SLz`X&&q'
    'cX+j&<>fMg_Weyyr6dO|G<Z4?*3m8%)yjybRl&i;zE-Zcf#RmdzH!#_rVCY5<7iLpn3BL$LIxN;fpKLdL!YgLzAkm3ug=h8>`RW`w1M8t()*~VEJj-g'
    'jAj<YhCXEwMAJsIsL!v0OE<~^X{#hNkeUVZi6{w#Lf>I%=2dI&%zlr#d3$*2@EXBt9H`R+{wug+J9<7&hXhGz-+HJiU6mM)a-v&`QBFP9XLyo@MFFvL'
    'bDP5W=pM<6e}JKqF+-@9Y*7I3uv=cmEgl%DX#$cNBXP|!kZJb-Rc!z9KPlMCAs%*z37QlUT0KIGg1C_7JkaAhICf;~|IkOw=p(nn?9#8aJ{pPJU0t>v'
    'OESqHFcbC&lR99g53bNS<B`6gl<)yqDl?Boe`9Cyzz+r}E&F~qTI|I#1@+NwnqwRj(Bpumf)2QKl)JeDZD72A$8;r>^X6$=f9M%VQcA{<T>`0{gR83e'
    'B@2*844tQz0QSy2IG&FIKlWe3wYrbC@0hNFalwm-Z|3mn43BX>ra&?psGIxMYI@w(r~4C#kSk#88{{?^HE2fQx9e33?vyD6ub&5p*Oc;DGSEW`?HbD?'
    'bTbQ$mC8O2dn{)J>ud`_jo0MkWNm$;`O1aq<ww06c13?24|Lb3Xg(4GMp~Xh<Lf&f+wDF<t9;Bt!<rLv+Yg<iK&gRVY<e6|B;9H(Iz^)GYPm__8Xx7-'
    '7(F^h*RT`V^wysJYi!ANhiw+(V}G6Cxm-{f9ydXbIqDAI?Cljc`=l;;Af|L<yvg`l<89#VDNi^wG5H-eQlmvYSfw8rU_EC3%iV3d&u#Fqa9qi=;5Mff'
    '(Ax?6O^WVaFCdGwBXr(x*r9zFUKr%ra?G<&S1G1<$OujWOEKYY?l=*V*BL?e7&KX8QI;OlhR)5LPg`#3Y89<c=B7qn!0iWolp)5VlV-de-uZm?{prxT'
    'plf#)1W2scDUPVT;gcI@j&ZzPxLob`+yT#l$*UB~g)|{lJrC6_Fzr8v(u))-GD5Li>r$vHg;J&0rs}rNi6ff<qzGUUu%SM#JvxM&<Y5yWx`;ZA6$F!|'
    'w@D6ROiK|3-Pbi`H_<W1&;@Syce5TA19OM5*~dzu=-Z*+99|Z!5pfG71T>|0?_lTA;Fs#RD=a9u+VrOl$fFOP>y{^448gC8&H9ZLqOErcpzsG++E~R|'
    'Ks&?~@hV01-XzHe<L8DbD+WZgW->k*CS2-E#y5f+4X(yyE+10GjXWpqo^6aIH&oW5Es7WQn<i<}Okl$2o7A$ca-GY)o0vO3gR+_H_-w&KP;y(zvtWY3'
    'bgp&01+&?%5}tsvISVEk%*;+xmQD*QRZY8*f^7yYW<qnEou9M^jZw@ziXfDkiAj677^Rv=!CB2tNV$fG80mYyDfu196dT`)T-5WRt_@~s+BRPCo9;v<'
    'W^IrXhJ^osd5!<_PaqyBY6^!I!!ejwf6JQPH0cO`7}j^x(xZ58-RoKfsOBux`BkfJ^UAUj*_F33T%jd-S-_am!UY87dtD<|a|xC#>f*6zQna?#x^#tc'
    '40iY>``{X@PVWogvd7M;<@ZI6ZFmoF$JE*-)*0-8)=sE7fPa?Qs?zxBdirT>Daa-kEQlPb^q6O)O6$p0XvpCt=MLTdcrDlch9cYG=nNkUY%c*v@dmqo'
    '1q_LC0yc6MA1OWdzhdu#hvL|8R+)BRAdBqCaALV;g4)W`quLUMdy24&s#Dl?eNo7_Sj$`Be}DPydqKi7LBPX%>{Db_2V|8;UZasQgPVhM-TUw)wRt$#'
    '{_he%d>O*{stE8326gYU{m5>05ulR3YoBF-ECOt(TEPL!Z@PAoO8?706%D>#!`-&0C}$=g1F{1V0ZnbYW>EQ^Ww-AK7CPO&Q&Y>~@GeA1qMd<&1;PYi'
    '?7Uzr+bV2QT(Y&Eiq?_TdKh+O42_yED-r}uisBieYUy_DzkdGw6$_|$KYse{@wczVuaDmzi~XlxzkdJi)3-nV`(uGUB%iPw)~DaTef^%|wSieTtQ7gy'
    'vHrds&~kM;EU{Q06>uKN@ItaM{1IzDk5OGMRw)`z`K79zRH7IVQH_4EWm&Kg<O!`TPddSHMN$!?k@0$Md#vk%H!GBecwqnF!H-GQiwICQru5!j&4Hs+'
    'Zwa|P5+8}BHG#Gc-Wnn}<D+3>C7rlXvmeb;V|$x!&Gk>$lyPv^U4W*lfR=IaXHoM=S_Pya{N(|1tD>B39s%+W|1)53Q#4=ypyp-@p4VPJCGEmV+M3IO'
    '<a~Ykb)R$MT3{IlC!&0eELQ3@vk)%gq^;^Yg6QEeY!BUribBnhI#UCZjshb_#4~GB`2!=HxiuN#$Y)oh^&78i5l(c9^4wYWxso`qOsM<_FOT5DT@WDz'
    '?~Vy`XV)kl2=j5u7l0UlDyl||3u1)``X^cDXr0FaX_I1FV~=>AysZMc8I#9Rqgmq0x$}{@O1_=c0?5*fHqh6}ymKKkc@!B<p-E6|%0iydxma2oh#C>v'
    'U@=)H5j2*v^8BG{lMVAV3$mYV!cM29&nPUuU;424EI{p^9|TazeuZG$Jgta6t`pi2rk<yj;h43_Ik92&c{&x2NShpy5Z9e2)!}rs$?4!cjXaqN$DvJ*'
    'gRh#AC9G5+Y*Rv8BO^=bLIm1|O9@1T2`ifcXni0Amo<$5vM<97PV1#Q(WVEQN=W$1cQXfDaMKP%&{DPeIGYds@%WHsWI{MGy0l7HGyqGM8@Lc&lXj6s'
    '#D33xsj31+CK}q3iX5rewKeysw66}6JonfL+scp0D<28%KA4-HZRp1&qg?(o!?gLyOf(m7WXGJz@wOm!*bp-p%yw=NF667V5y{i%?Zat5Y<hBctcR<a'
    '?UMePX;{CC4@L@*JtE4#H4XQQvPa~SuPgt~G%Pu%hb0BcErpWN6yjr(5I=i>T-QSa8lQ#)Do?Nd6VR}Ln>`rmsAe)8Eu;)X5;MnyO%fL^mG&$c_?euE'
    ')*H(QNj{Cm35>OklJv%;GqlE;%#k(|ex7_^WHG!QQ&c)iN*~iE8=k{)fnFv0i$|fG7mUL8+yiRKna<gO)JuSrmv<VIT*3|3yR)%aGjBsJWgriF9(!1Z'
    '5kiv(P;CG()9}j(AzuiCVHy|a+k(tJcmA<X_40iDrA2e$zJNB)QDKPu32%MqB2dN2z5LndSU~HuBwLoc)7)Ny$&-42my~g8FH_{!t8K!1O{8ol%W_=T'
    'U6APVs?GdYB^QmMaW>UVR4I7@Bp2puQhIOV!(xabMd~b#E(1;G&X0&iN^I&t9F;?HMDeC(0c;CH28dFmMjtIxB54BA$XaoBuw4C|$!d9yVV{LjL0~M6'
    'nc)ce<6(>PSlTMhXpX3Jr$c#rq2P5-cPfx)E`hYI`<T3m!=4@4b5~H>etZN`%@lbGa9MkbtRgO8Cc1Z$lR?(w8QyAhhtR5P+v8?~%UKDTFtr?B<L$BB'
    'sW==ftI3Fr?M!moge8m9iK86M2~FDs*VVk*lo@2>jqXB^EB`n>4W6+KT`NjQ$0Z5f?laZ1@3FGMz<7H&g*3O?*kv9q)#8FzMA?2gj5vDX15EB1D(zQX'
    '){*fvN$`n4vQMkheiI~@Plyc?fjEEqlqsXpXR2ETgfs85vWevwdf&`shWio@Ir5--nljPbxsb!$=)Lk-l_FCI%9Sw6k$No>&U=(T!4;GyEF?Mls`9E}'
    '#pEHaa{il@3biVP_){gEa%f()thu$gWn<?CPz4>|WNJEw@XT|&T??<W*G5jHL0J18M%WLBcgt?qFVGP*b9QV0(rYd*Ft+ndVP?)jt-M~({>yY;8LYE2'
    '^a25lYc91j2YLa&s4E??tEAfQbjCT((Rj%LIDGQ>bPH|<nti!dqhz2|ldbc_+3nJ0cv`CX&iH|!AGfxd%}8ID$#uO5z&5c#Z!#x9WLajnUu|dSAp8RW'
    '5Xq6Vg!x0Jg=h3YspqN^62zIarG;nv4cBPR&Vh3kH>bnRFgp{^_?t>3_SF&j%zSL-Gt)P71KA9RBevC<8HAM4AW{*)k<JakvwaZr%b5jrAm^6UyYLZ6'
    '%*_P5Pj|~<{&-9OngC34ts>&N*?8Mx%W`?#+9ay=lf9T)Il%@4Ffxg`i+$H`_rtMx@!Pk@&ySx!ef#=`|Md9s?Aj>*Pyul2={N^EH&5lComS13X!u`4'
    '_ODaRHQmFTnt>|<UXE387PtkAB*n8M%R*J62@-lalYzPz78ngRaPD;u#Hl?9Vp16=BF-I;N;<6vsjH~U)iYm}B9PTN`6!~#Neq>g(XivvbP?rq3RlIA'
    '_83HrH+DRL3~3u?Kuk9;GDU-V_64q#h=yb43Uox6IR(0MA_9gDO@*N0z!_T2?<tvM|465%7u^z`xzd$!j2|IJ8lL-VZY8TS5GCi&k-@!8iDdMtxmqNa'
    '{DPIQe8!(Gjgl>7JM1pOYMksSDpe<*v(T2p%v`8O1>%fbU<8m`&BocA5Fm01#Dve@i^hq92r?&Tju7X(J(Z6+tGIi4#-ss2hx3#2cdb(OOeK}k!sL{>'
    's+A_71hFC^%+Q7H9M_zDb}!mCEaUjyz{HE?sHq!Kfbdy|5WU1M>S;3fq1Ns-#ZQ^Ef7x<4?Yf(lYdDFmV7jY$vwE#nH1KK;?WU(rGOKx<)-!2alxFex'
    '%TJ%a{rpSO{qpmd$IrzM{`^}?q}o74Ps(Qu{);KHy$&q3t^zfkmB;U2i{01HKmYpp>!;6S7|+a75TgPJ>HVL+{q5(^pFjQfn2ZDN@N^(pN}w!lP8LP@'
    ')=`8nrd;+Yby@QpMhM3dzD#X{vWU9MQM;t!Jh=_XB65=>->hyrKLSzaFtiL9dKSZvBM)N3y9lJsfnf7d{PR%w;b=1g(%7G`R8F36xtPuBu-YzrycVl|'
    'fgK=0e4_LkxQpwQxOVD%&3Ez)abY5AY%3h<I2&E@mod3`hspGt;gA0^Jn5!*ap>`N4o=J(Gf>*H^*v=Y%Q^eiX7zEocFjdjyJh!*BPskC)~hSbpyY}W'
    '-xEaP!T=GmIgbc>>vF(E_w$g_g{oDpB71*;baVv0>0sW&7GwZ3+G*>+Ci79vu>`}iyTXLeYfOnje($m@vHId^y`=tF>lOYlPoXJCTRYiU*gsMbJa^FD'
    '?YoDNw!a)Um>G%Pu-qO}?)&#h9%?B`sVrOHgGh7c6KeD;n@dGAyY8`9NT~mxAXJU-L&Wt#Y`N3ZaWSlY(jUfIsH!P$HVJuX#no`?w0qMN0C=yb#H|Zf'
    '!T{aK{Owv$;#Y4r1=ds@5eny#uUi(7?5ugd0}h`jy%?^j;@pefe%QS{ycJTGZ(tz9`ie$2B)F~fn8JYCOptEq2zrHmw@K)O0cv;vM?g}#-~|doqi}Sx'
    '@$3dXx7gD79Sj0wIP8>XqoMt5wS3d0c%j#n(<$;Z{&&KrgTT;4LGwPm+uaVNw+zQyUIab61kxpM$f{vW+1!^ltNVV5g#ti0n*ETOY`XMWSmdcs9jd$r'
    'O_?r$X?!O$L(bf@3+g({WAU1|eo@;orqWdyNj>SE;!hz}6;bUQk&=fNXlCrmKJNSO=7jwW?X5#mosSAu;d$2<@~x)qh_?x*ab+J{!*#KIndGA{EYet<'
    'YBC}t%P=CHQrnD=aLlSVDg_AL*?7T60KP<7453p9=W0iYDuwX1JR*cDMR4t4U=3=OBG`8{EJB@<*Cvop7J+SaPfbeXeD7KcW%5KD?C!iBFJOvOmA*rp'
    'X~kshu1;KpAks~&gbWZeg@FGIqkNlyRZq@CUHb@24wgQ9X`;4ibZ^|X5y>ghJ^=w#%ft;sfAHNFUHKQ^DjhTF@aW^of5~@2Oy31eA#+PmjU%?$T7X3}'
    'Xtcfw7F|)@yc2qg<~#H#$#sAdw2}8d6ab<EMHBpbitj^&RoR&4JpQuVoUW<QvO|=a5&Pg@g{?ATL}4s<x0iV`Kv`wr+%t=CO%{%x16WcJ8m5rBIlI_h'
    'xG#0$?}k%wI`)T+Kd0_AOpsNXL3Vv6lq|J*204ESl*quk&^FklsvMu$1*p2W!w;+d^8=O{WEV_~$l@8~nE_|+K<gxPCD`=j3Oug|Zrg1Itdy1MwhB!Z'
    '+xQ6yC$B8r?FRmn@J~z+3LDVm8&PHNhMGS17*y^C!)0*dTqQDKA(SlxK|&^tshuK8m-qp<xRuY`vVm;@SJ=Fmi%VpNTW3B}h5mge!z#Aj9oluqfVR-p'
    's9UzhGrn4ptW|Y>(4kQ?qHj-&@%bxp&~MC?TYx;|dj{5R64p_OfBZMR;;65<Q=P7cgU)Q9@l}i@X9on%5pd~Dr=S*d%8iT<Yt*1ZR^uYo*a)~^^!AaD'
    'F;kFGQdNvIt40G^^sDXN3d>s1DBadq;+@D0%w<2;iesdn!5kuK#^cHit_kf9)#yS6sDm>#j6OwcScy04DL33I^g&eZ9PD%@pFwpUk>MzfQAZe~<~Dxh'
    '*Q;Z1;bjio$9`&~btD@OposL8tV@@I+FR~GVsOGuD@?-Ws<P4nDk^_@4$DHB1WGuCL~B#eUeXXFO8SX@BU2a`gpCRC8MjZk-*zO}4r~meT-Up{Dx>Za'
    'E^2<?T_0n~SL&2Lbhxamc}ZkD9KudAYPwaGc9#Lxm%>fB;b`~FX3@I>Yn(2Qadl|mSn2BQvGNOUWkz?t{4Hz5WVFWB>ea@`D6#Q^(~vxWp)XHm5^g$W'
    '+52I!9_|qfHf+k7Lz5oFnu!{XKyW~)KqMfI4j_7}5?X4v_EoAn*FGb=vUd06U?u}d#eI0Vnn0gsIayj+y00fncCCGL&RRZi38B+C);`h=3QZ_%V#>^Q'
    '?XE#vb+f`)zXZnRs8rU(=K{F`dru+KovEFXr)navqNgg6X8HiuwX@|QSscrrd~P!FV3}cO;`ks~AE-lV?0Kjs=b?5bDiJ455wz59KkLbU*1jX6&4x)s'
    '>BqB}oTj!7vlkN$*+-99m|&*D8R2A7CAX2jD4LCt8v%x}U@TO-pvahlSZLK`u-Z^yU5fKqSoHSEBPG*pv@S!3myv?R6(0*Y^2KV~ZK%-5AyRBTpKGTA'
    'D*Dk~Wi_sWHn#9!!Sx;4qJym_yca65w16Vy>$H`PIAiN0;Hi!8@bc#E=4}*Y?{Iy2D!^6G-TC+%?&`O-mMN$b0d%yYV*oyPN~h#a70BE3d_7vnxFWFI'
    '7m#Zw8&fO4l!~_g{FHMazyK-7^<;AGqSF788&3zVEO>10`n#?mXUUNYNGCEmfD2Qm>)N!w{R7wEj{q$Y*C_M|$*pOHlf`Nq;F^Ejw^+t|?~U>>QPaE='
    'Z>AD+TRR&fdX2~F2K{cpMprlMZsX;slvB2~d$F2P^S6Y2ygZG?YEmquH*86pCQZ!a*|DYkul4m7F?m@}OA}?ZS*vAO69TPVRlZZnvZ4RNc>+t(mKO<G'
    'PL6p$INUdr7faYWOm@uXk}+E=t%a<oh$>oz>JmikJe`8L(mWX1D*nWL*?5+s85e0PL@oaWnsXA{$%-<wfT8>kXXajt9kHMA0g?HDt3HYASzofABTJ_Z'
    'o!|o&J|w3f@;%eqrWxOL{?IL80lT=SEQ-lT@A`qn4E}<6(_}v2+t@Um_4wyLz~+4m4_lG>fG_?5AE^KA2jYVbA5c^L+C>GS0e;>zv0q^|Bq}rlBfeRY'
    't7(gI%=9<pwi=|?%mi8;eW1xnu7_vQo#QeUYOy@73Fs3E<M}BtBh{(ci3L*K{j%LfqQ)C8?dyog&>cs9Z_PU~F-!>&!mxFtL_VH>x_eWbgb>o_;%JW3'
    '_w%(dSt1{RZJ5*M6uW37ISl2?{h`Q!Q()I42!U7K67OcQTO9^$a9X1mv%r@2V+f<$u%xZefLxNZ4}%hNfwyr8FIiV4mcCDbeR0dCt;+ln!)!Ajw|V|E'
    '^p*6aiOSc$Y}#rfY@dg07&1L$l}{jVxLjx&dV9e*lNTF6-F?Ht#`?VExa&9+(v>)j`B4c{-Sik#l?>?SIlJ~|(Qi>bvaFhF0_IH5fLr^sp(~O1p+#?i'
    '6Q0p-Ycd9SxN$ywPc-wk?BeFM+V59b<L$#63@WvP5V|uGyw(MCD$Xi=Mnj4Wq#<K!J>22)TEW>giH{#{cNu-9dj2cR)2=5UNq|h$8^9D1W_<ysAMOZN'
    'q_D04iO{A!7vmB8l`J!N5f~J8QC9OYF8IHi@R{cF&vfB4?kjr!E28yZVWW|~iHH0FNvM^-6jeL!y|wH%{efB$T&>0}>n9+bT1kP9rpktl6miH1xumF<'
    '?lF!aw_r`}dzEMs+iy<&e%r0DMx`9a+9E4+=8a4bP~*eV8?FTM99R(jJ!X1(+R$6ZkUr+aQLjIJ{`v34)z6Q`=JCs8vHS}E>hsrcA*izjC6G#Sv(}~X'
    'PaIcimw9`$T9XHf27}r{N@^oIk+c#~s~-3-O3v<gAXn{3dg@YITj!2w&6M`^p*Vavt=B1;HPxt6XfklFS5&BS*WY8r^wsqb34ySZD^0H&wKf`x?zCSG'
    '>(0*>JH%`{7WcYCie!t|5mJ32lC5}|Ks74N!o-uO>Ql$nJ=e@7TiANVg4X-fRl0_a5gw*$b1;0#57_7HZkUY1R;)nsy3CO;Z%@k)4+&XogI$PB<j8LO'
    ')qc?xZ?GlHwp;Y+hg<FTL>QI*@b!Lmbv600^(qrt_QM+xS30y9JgoMUk2Ybk>MA>1U@}|sJ;1+$Ijqz3-nLB`fXqRxd#r-Ly2r5g!!g+};fk6Py~@!a'
    'R$I7g!(n;a4=J$-AD7uh+jS3Ui3aJh=#MG2sr>Be1@5;Stn!-NFIC#=ry!<kx-M%Bf1-Q@I7YfGmYeXfrL@X_gPqeE+Hr+tM;M(4QvR}pYhVoNCtY}e'
    '!Vt!WSSlB;#h~17)4eOO_Zb>ee0VtkVvXGwzdn9@nWAYC(AsCvNF_WJW){qN<|yfqUoy0Lv|0u<DnbqGQ@B#$SJZ&JqvXr3Xrlf}j=X6Dxtb%lO-Y2x'
    '0JK#A&?x|_g!==9F4MBSBhyl59Im_=!I;Zna>H?@#t5fMao`_T`_qyticocCz`;C8j19%>og0{p=2u@fFp^2T9k$oYLMjy?n?DbkowqInW;@wOt~!_{'
    'G_@zRhU$or-n(b#Q-!;x#q$u^VbtM<X?bSfXvbVji_Bn}HY^0$dXGIa7BnF6#T@_Mw(VpVQ%4;RfyJCSn14H@Bu!vv^zvAet6@Qo#%;>qwQBs!$lf0~'
    'eb|VR1t_~aswtM6l(3b7(4>Ui86-z%3-mBumC~^mnx#`=HHb4{1{C&HqHKn34`(wJjUNgnCkVBFi_3=Oup1B+sVQc+{7^g<mFrhSGA+dlJvftg3gsVW'
    'fHFC#nFmML{yoMv>N(2W;XMtY-}i9HmLw~VL5#DZDhfA`Th|X;Y{_{jzW@I1x1WFdN`Sv7nNb!*2)FZaG<oMSEauo3%YF~M6)msb`mVc!L8NCBYh4Cn'
    'lO?{!5WJL5)&8TJzW<>=7OwT)zr)oYuDdP%36tY&Yq-h$m6rlbe(}$v@F^Z+RHbycQB-lBg3a8)leIdcsTI_$<qE@>3t7&rrHz5%0THXR6|dq7H=>4j'
    'k_=wMeuKq$_jpBkgHkx{PpVGK@ggSSF~H&cxGH~8bR2DBTN@5;6r!QN0y+Ly2`^rix-~O{fm?D-?-WORpBa+yf>Wv6>Nz+cw{{M(3X7j+YkEX`Gm0ZH'
    '>ksGQ!rkUXrF6~4Q0JD;T5S1P=rhM6*2ZUVY$b!*tJC_$VWKhA{48bAgE{;^Q-trLLD>QlJg9bNP}jq7<mwh@hNOaE%;^NvAw@Xz12WN>*O~qEAm_GD'
    '2dOU(0Gw`JNVGcquv#7S7uxpxAZ8=PJhPdC%`ca311pWnJx)g$ifPX;ni0`)(D{mvm%vmi%u~+KGs2pmpf&B&m`XWwvI0ZB0x80C!vw0Td~+^i0)LJU'
    'R%{r4s(yc4T<w;nXi^+6l9gQQBYW6KSZ5|!d1=iS+s_>MG1lzDH#+D7v<FTi!1$l#**$Yn6^eAzk%<q9;cv$tge;px7n=Gdk8)+L?u;|GJ4QB9itCR7'
    'F+k^-#S=Q<ndisFS@kK13@p>Aiymf-mcm@Y%ss9}P`R{eE`kH08*T0L8&Hd&q!=-G&atoE`?JTXmeCzfOILXuzIya8HXSv#lGAC*Ch+vg-*s-9$9{3z'
    'V=D@rX9^cxWHEPZB81$KlnLBUshlL!p&9E8ORseq1exlxvw)G%90)zXeJp|6+06+wQAqWf9n@YKE14nZps=Cua2Ce86KJBfIWy6@paJS~ecCThl<Sha'
    '2(=0Ks;<wjW#ko@NE9h*5Lb<?W;uR@csT8k<OMso9Bp9Qa!x*sw5^#F*hFGx6mBq`GKs+J)3&3;9_v}l&bJB5pl;5p4OR?xs4fA22Dg<F?z#JW(I{IL'
    'PNkeXl~W{I!|aB*zQ!~>d1iQv0j&YHFX)*o*apXpk#jO=0(NG2ZP@5WJv%&K3?a98jcYKBLsRA1p+!KmgKJ9R2NzWt9E~j*P!>i+Oe^gx5(F9Yw9now'
    '3~u$}UQ4$Fqil-no<zHw9?Q<#tlIf00g5RDN|QwiYTy|de?LV_uAwQ6RZh=f1a}Q5wNEizo0NGJS+|jIpiO-Pw%s4Q2zF?(8ra%&NMM>WSiC5-WgZ50'
    'Z^!=S35fEGE1GgFl?`3Ui1Sh%KA}bOb_yj_@MBRhU2ZIj#uAu(f}w(PP={I=+*BQ9x+98+H^nOfPxq_hO|kg;`TL`-X!H9oZxb}F{0mem+s_nE{sRWK'
    'U^qDy$-*e3o{#h^YZZFMh&f23g1409@#E=ua?Kr4uIR2{Ziy`n$aEdZYDY%4@pDv02dc(V+0c=5bU{Xt<lPNcWb2Vh2&!Ok)g)e8_wx_1R<A6{tCS|#'
    '1_gJX@6dn@y6N4J{xnzu1!Y0Qm0^?YaM!!KL@AB8<=RG$j^?gEUD<j_DT!Bb1Q#>9Au$LIgn#_!ZQ-n_$g;+iruD2&vfm}1Kb`;WYS_L?siX>2bwb5f'
    '>MWT*LXv$>F`3VF$Wr=KB-zarQ@U~t8N#YG(I=H(8M02ToJAc8Q@SepU^$^14K^k9Y|;kek4fXvCX}`f6j72a#+28m?L(j2cTj5lC>7CKkA|(GmWx#Y'
    'BF`IuC~|6C`PyqMQX4*|W?>+$WFW1Ht#ntbO-wuUO>Y_iLb}gbeQAj3@T?_%WYL(~oikN7QG40o$*`Kp7XCk>t-OO(iKsjMlooh9O#gJ80C*26u+w@T'
    'V9}As_lhD#yG<HpAe!3y$Q0LkngLpZl(G7l4x{##ETBF!57_t3uDjW;5_(&;n%D5$*KM>5tH3EpU3w5&*BKBpGbOaGKctG22xn+l>+69t^H=a-rXc<R'
    'DXv#G19H8_o^dw`d0TodM~(Uvr{wVY6{ePU*qA%P5JHA|n6nr>g+mJ96&l6Gz-Iv1D3TD%=qglaw$U3&6M7}W4z_1L{=;p5NFszptHTKvGW+qv>Sz__'
    'ZnAl8FvbG8K2L5-Gb9MD35<tMg602tv!Z+NPIvQCK48}V>8p&%U<Rqr0X`Kvk0w!S3nSF?B`5M#CK$q7f+TwagEZ>SRL5>=0N~c12}1CnN)sgX;J9`2'
    'xx4Y?g&kvvGDrm6Up2*W>W00wc99Yk88*yt`{)FPWQ(&9%DW9EXXRyri1OxEDYVzUBxkxzJ+Kea{#v~}1qm`hRqAI!yxAdvsZ|(@ahq}|))*mraEf3b'
    'A_G=b_7o)i&p%@iag@ks&M;t`Fho<I$L77i7};NW$&D`<nYBI8gKAKNNQQwIzbB8M)X03OmPCTpc-yOFU&CmWp_%=`kI`W^vylE(14wlaGA!xypeAqx'
    'GCR+L(&?q3>1RRPHf%r1Xp3P_$AxWl4>lU)O`)pd*PlQA{qZ}ti!W0PtpxrM0qFn~#&Q8hxKr#!bc5OV8x(7ScLyhRhTI!ACHLmbkjGBx41y1FkZ}Wt'
    'k%>ksYt&oVYGw(BP+^-ybpb{g(!>UZ&EUd87w7D4)b2AO4qiCPIV&0B+tocbFz8|fC~ep5+a&FSWEa{dRcv8}UrmhCuOh(kOu<4=mcFjTEfmc~V@*?5'
    'X9uiQ&_uW$Mb2b4WBZ-66#;5cRk#{|*5rN0{5ucrstsriK4WBSBtp|jx1zGqu}yyecK^0`0w%XR;BO~z-oDnjRvsC=UBVr-@PT5owbrmBJnZOx+NV#m'
    'mfr9nTWj;8A5T?7X=@4D-*Rxv{zxh-g`TfICcjrY!Y1JZx!(3j)qdF@MhO-PnfBfkAmggV$JKE-EQeh`z|5~#`%Q+*hZjj})MULi3ZNxn491D<1x;}8'
    'hldG@_WA0H7*pf}JuPp0dX+(5y4m*t27kW2Dogo}C~cA`mPSv+Fm1xK&vskbf2;v|HQW@(?k4$l<NcGMfuD7JfMWPQl9A8cCf)dm2v9Z7xjwGxVPARa'
    '@(gC{W7QQGUR!rNET~2PJQOMqAc`4%ubTVzwJFxW|NV9G{VQO7eO-L{`nR`*>bl~Wuf@;5{P_88iXC)Jc4t~mm}*Gc8XJYI?qz!q7YLw>f3w)d?&r^s'
    '#b3Vu_VpJ+{vF1ig8Nq)zzrR-e};1br;Lo4{jy&YLr~^kRIkmwEQ<qU%QGVm3N1%MFS`Zk-qbq)L#_>g5frA`e1-vEFi!6mR$PB1j1#XO24sV#Q<osN'
    '*5$pZ@8g6tM)0o*Am9n``vSvy-(f^IP%3pVn(+?7a{le-zns}2Q~JV9T62sC=N7=(=ITr@)~Aom?%rU9tgGQAFi-Q;$S0xEX)N&e$TRoGls+UuDblfl'
    '*^prg6zBU!-nX>z-@g7Ik74)anGyTo1lbT=x5qk(ODvq~gl=aM-mE}jhJUYN*Tj9Ck;hO$8EJ^wcUaxqTa)kMuQ(@{X@YjxaNn90Wo%liL$^M5Sk6X?'
    '`}SU$gjxg!*U-vuKykrXp96BDE9*q2EMK<qJrs-6H96_m^c@tZVNmyRZ7@dm(T}pRH8mKS-b3uSf|+&W!0iY7DGAG??xjj0y<^#WRqRw9Vz<OLQt&TC'
    'b1&M-u{ELz%S%><We?AEp!YK>i9B6t>C=yZ9QT}<+8VW8q<Y}caXj*L?o;ZZ=`|!ETWbC!uEGuk^d6NpHy%*|qex+zaN0O&j!v7yd9X3m8h2!Rz+}+7'
    'Id0i<6Zb{lV+V%EO-+|ksc0R@#I+qZLqf|Ff};*sF^S=p*rL!Q$Y>)RvAZT5!NA9j1q~1w?}ffk51SU@7)w?cj$uFXTHn_xi3;aavWjpdjWM2qy+k!!'
    'oe6y%4ig17**Un2<Hie*iJYD#-9{@=1!{Dn1!7&|OC-{JV0;dW<VKvH%~(2tU%a1?$>^Il5!ZjpEx8iY?J~hI(6Hzb5!%!pE#SUR51i3qV^~0dW48Ci'
    'xW;^LV4dKlX;Shw95qh<;bkHLUBWF-Z&?*Icuw->>#?KauT187Fh27sxehB!f^x*m$HeMhWVW@c!b(x*O+>%zon~Q|%v95*u7aM{*ld}Dw?aL@vQfF~'
    '2F`Q3yLC`(8j_&SJVSN>RWtd1ts?f)&fINRD6y%h^g66Ry^G1swqK4fwv8<SW&?aT-H_Tu2f*`O@9%bsq{|<x?-iW;bXS5WAEeZokmQTyQ`1&nA``ZL'
    ')edcy__zHr_3yy%>pJ{ihTrdpP3Fgq4nL*<x%b0*+wY55D>~jLUU^xSWkiQdwqM;1YY=DC--%!#C+h@AzsG&IMv)Km(?JPAD^W`2Bawg#XLEfz^q_R6'
    'uTxd3U_e)#7yNNZf1`uHOLNsIPJCT-hZRP2r@uErLA2Ry2Ry?V5=$|I8LhGkN!~CJdZXZkq<%v9Srhh>GRt5s7N=#37E-wtyFQn&Xdo>gSE&~Xek&qK'
    '#ro-fNOF(R_7I#j>?fitHuf!9;pl*`^g1=zRWj1(bd(m(mUpZgc-wDQ+f_PgCBvQKjtQ8G))oJJ>bKkUx1J#a6~*!mo5ddM?)Z>RcV==y0s!DSA^_Cq'
    '0dgCun!wgb(#3VcQZQ(oPtU0?!)sI5BagAoia+$j={}{Bf#=p?E0AcvI-~>Ap2NX}OgteDIZl2a!b&8EHhErrxLxi0>uYcKO$b9ohHP!#?R%QZeszO+'
    '|A`1x)m88uayZ_wCcXCdM1}R%o5@jayohetSBW{RWC(&?aVec;`9$_nIE%-A3{aOroffpu^gNDj8TGp+r3%`kplfQWKtIOve_S?~AXXV_;a@1l!P}NP'
    'xj7la@RV>ci6uOC$K#MC2(bX{8ny&nqk-F1*|?$%aj>ShA{wR*J9vg#2H#do&+z!ubPCW#w^<CC6>hW$uQ*v~_-ohQ61cW~E3GJy##5GQK$UEa2}T|{'
    'x*zVfv>mMu*g7{kWKGiqPRu=GY%&QvXz=jn48^%@P1%-_TiaEFLsK)})6)u*t*?8EZkRZBWm{Fzq0u(&bkC*4Xz(zPeUFjTc1CZC!>8ZSK3&)zm5CJ4'
    ')@67SOMNEj0vON{#Ff*&PY2aRH<nmqFbgu6wPvfTd=kx*P_I?7OpN7iE5g%c#=h8HhxTy`0(bbSc>;;^OB9ng=Ok`h6=W7lxWLrRs_0<J`{j@+6_e73'
    'Fb3|el6QlXd2xGM-3{&pMq^7HAQ)xV1iMX@0S)zX157o65kXMql2CIt%iU?WdpMIcpvu6DfJdK;F}^;n*Uv^kV?<@(IL*1l6zbOSFv>@m2aPz9($5X@'
    '`LJWaiMW$8&mmsmM^a_&qqS8zo01i!-MXz2vN$DBO%y;<&Lz;ix&jsU)tRtpc$$K3TbO2av-}&L+|VYypGt&KLafL(H=6gS?ajJ7J02NogeVndAEMV`'
    '`Qg|BXD3O7y$SOkB_(xaU?(F6Iw5R5mqoW(oOtuX9g=1i%8QbU<9@6^WB6aeibPEcOabY7b;WAgpM+|gVCHYQ<_xu>*<(ubl8WUOhtrNCO@|)}(4Kw('
    'PXGP!GoU3{N`xtD6G_Xj@DaNK-UTc7Dxont(5UgzS+7qRP-~?t<jT?{9#EhUF9R1$n9w_+aTcKoDz7ec2;o}hAwVj2lhCDPkJ<7BjNz{T(5HcGN{2Wa'
    '_%#yVcBqT+kGu58etz&{T+<)A^@ko5Tl!o2V?P!6F{k(+jsqFHZnmdQ@q)r_itgsry-Y}z>u!f#0i+)k9w@jfD6`)e0v4fLs(RL8fHCUP2uK?1w|(KU'
    'HmTIFtecN1SZ4(Zh1y_rXdyVHKX>(`ah#)TwA!*DOg6=o%7-YZwr1b^forBN_UVdi5p7TDmFQ*nvD<?vwAtZ@(qU;IldRZ?VL#AqKin3pn|`<K(=nN#'
    'P-`x(Mv|x3^(YXBGE!V*asp5VRa>)iFb81clT&&bf;dUvnx&rvYyd%EvASx`5l_S7lx_lF<%B&DF&7vzW9hrIA2%W5M04O|Mra@QPGTSSUR*Ea+vb(7'
    '(+g~5ycgqO`~K`E7!_?oU1lk>t8Y}a`nAXsW@fIUW#^#X2~}pjP31LlT{dy<$lYq46f}faQpi-Z0fZ{t3QEXG5v(SHg*0>IUbWHOAs{7p%3hAT(mppq'
    '3!UU9I4LY8@x}xoX~ra;pQkiF(?oOQAxq<;?R<v7zvT=#s%nL+9SQP_3<eCEVDV&9`MQ?I&x>;|7W>nCcrS<jAO8`95kbFux3DqX%RLsW18K1`px@L_'
    '7!5S_7k_*F@#n9<kYhjnRuw))_x1G0e_1WDu2Z-8<G-%)&b`8XhM~Z`1`twT70Ns*1sn2OPCoVJ_fLQORD6Cceti7Pr@ud@->M_eH2cnMgAX+50)uAP'
    '20D<giWjDNlJZH4AHTe`K4~TApO{)Hf0p^(je;Kb{qY2Ago!m2Wuoy52fuNLq7_eCJ~i-!za{tocklWCo_;30=~7qGXI2|L71%hO9yHeE3!`3}CtW>F'
    'go6-81URkV<K;Wovy@z+C|C%sbG!QG>)(F<^7E%(e*4Pt8ZWGK9R{6zDFp6V86!k1;80KdWe=-zGZf8}h6B`0h8kjc0Dv@g>Iit#VT>-wGp`od{2GLB'
    'C7<9g(ehx6zkT}l)2HVaRC^PemiADvS>p<>FTCDY09!sy(%9kYR$*F|H}9?1fBYwsPeGnp!zk~D^$`mO$m$7Q%1`j0OO`VImJT<vvRU`G0M7nIo;27e'
    'b-^}`_UZBa%k(>SSPXjnJGgVS#-NhJg@=WKR|A8?eSn3~a2+)*`e>>D`0vxTHs19O@a_Uq#10jrEHDV{=8ykth4vd!;p=EnOAy6m@hk6w(6VM<84c?d'
    'hN+=o{;H^+YCgCrsrN=`R-PQ`4o=k;2y0l%CL>SRRTXxebtO0agnzt3i%17AJho0lbJ*jTzx@1I{Pd~#@$2WW-#}e?k}x(MEm``hI&3)YiI2}YbRff{'
    '@r~U6S|O6C&-20Y)~BCp11BK+yp7kvgsBxZg?Y!uEb3zP9_#x(0i)OIIqv&bCCt95T#<O875f&gxf<Rhcu-bkcjgu{Aed6|V&MFTT>=xlfxsD5co6Tf'
    '%Gk{+SH(}+))376$w-3^P|T3ncMl8vDpdiF+D)uiQr`9R9KUI5FY|NNfGnbIxlC~T!#TKCy7KF$Y$7;YT$>eNY$*uLd3C*B-CJ0lN*PTn!Xpt<O>;mn'
    '4p3%>bvMQE<!CfV?NDuJ4AmUg6xtun<RAYJ$vZtLb7PcNc+0akWFd9BcelWR%S7#NrH@mvy3y#Vz#y-bz-J`@L6s6<O|_K3$6y0Nof6=~k`O5JkQf%+'
    'qB^X$5RC8c5iLj2Ei0+s5pjy<JtvyGKSqe?5a4bZ`|iN?ToY>DM$3Z>OUDD-`{MSRii|77Vm^I$r3b$kkDrUjzdwFUu$rK(2ovRI?pJsrZw69#Y{q1!'
    'rZoU~3$bRq^%(#ZCdTx-6e_&xGIBjF;CyVU+B@7MN`bIGlLs%2SE-Dk%MjfRtR+-TU6EHQ`E3YphTpc)Z}+Q*<fqEV=4oZy`(k-&AOPukoBr4r7KI;I'
    '{>M2BinbMD?2ns~7qWvB34ece95$#%{Q$GRdX<oyU=CmdlP9NRd|KY7vo^jBkKzvL1Mbsq)qQv%PCH=zzLX@)o?BoO%_|~GN95iS`oL&JWUixXYVfy1'
    'x3{c)5Vc7#Y9ErN*=qq0E49lytnME|!PT^!JScBIph_$5A7@{dNW2V}wlz<to*bbjVm__}D`D*4_HTD@K~+a}KuzdENU~OfkKPrQc3PemBuHar+H|~a'
    '#0T1G(zktv#zeL=o@MLV((gA|B=a^Qld1~ul%9X5xT6=nQ}1WrDS$I&fHQM&QKMr-T&~U_hCn`e&ojLMUTxqayh@3}%OY3;w7myhkF`uK57eKlK&2w8'
    'TP|d_h1FlIk_Z+!i}9<0f4l<i1K0&{>gBP=bQO`B2jJB-;aipPyxy^T%=IFF<<txVzmLub=iugOrx!tl@Gs515we6+V+>H%fuC3P9OX<^5l|vs2S+&C'
    'oF6!>fstbf3kb(ph8FKX%8yU__g{W0`mK~@u}M%u=?~41UDh%6y1JJoDN?VT2>xo`H>GU?nni79Yh$8c3M+9eKVTJ<S5a%eSznBn(KDzi(uD({gJ)<l'
    'kejM0e){?Qr$7G3*YGlrk(3T^O(oA&i+V*pY(l%3b8}ZZJW8xIW!+a3L}^F^I^)j*#o5P-Y2Nb1YKuW#be~Qe*ZpG}6a+u_21CuB{b^dx6>$6eRZ16b'
    'z0L2O(Cs=5FI^=`^*BO*2soE+xj%hGM_@;>i&#AW%j0j4#or#k{QC2^uZeC9^u7+h=Ac&c-9CVbv!W4j8=x1D-+ufGvKyYx*Tw(-`O}a8`bb>^Qn<<o'
    '##)Pc?06&B)?I%G?-eLuQodGC`t?NUfdk_o1g<pbUaZiOE^;*aK%f+*cFl>}w$w8yT)yOmd=g)NOQ+HV&cRv2X-&XaNpR*Xt`GP$tqutHvy8x+92kei'
    'UW8_BofGU$$7KgG3?f7spgp6o`u{^DD;(J3r;JSSg69=Rr=4FCYx2B0bn6qRGEn#)R!)+yb@Ao5?{Co*blPL)tuEG|e)^t*dIu$dRwGb4W9K1F7Ba{r'
    'HC?p;85AC8Q$SqtS|+$Mi|dn_bWJBN;06oD>40Wh#55O;)dvUy=6FrXLBC(3r*PRV_A@I`Wl_TmT9dF8ufUEpVAy^Ty7uQli~Y?a0X8xMu1COZL*4!H'
    'Uza_IFn{6%nBWMhNrViT-bPJD;IIC~W>vxQt*e@I<<?v~ZN>3P2Y;(l%CIF|W(;*Rj6!!uXEC6GbJ5wp6W%U(z~|!g?>~b6`T5gN#fx9Zw%3=ygqq}A'
    '(uaI%t_i%rS`pE=mee?R3paoG@X8h`gGWD!9Yw!kV9>E?{0jGbf9&?I(b@)s&CtvrCmv4BSFa0I@bD9-g|3vZ8K&7XOR8%DVrt)IZDN1<_W1X&-($qY'
    '4qH^fSp|9=(ofaFfMPpm*z*v2lF%0uD|Z2E#gow0lb9$vRcQU7jg9d}&-O?3)~ycl3Ufc1LJ1h$lb!+vMcOEqtB(BmQyWSGdf)A@`fX9e+tWUY35Tw#'
    'VBF`ds;JP@RCli`Rfj&BjV*qHYSst%*bzNPz`5`GBrgPB+bZx_iY>uYhOHM@*j{1Zp-7Z$b8VX|V53jYIJWu-h_(HqqvmSe60>H<WZL&?lk5z&S-sX4'
    'a@+mT^%z1hZm7E@uhWxAz$)M3&z29E19P^KS)@?FLpS6Gv<U__Kxf%q|M8y+&cLL?d=rI+*Nv8*Bbi(n2ZwL?s!;r=o6{B-i2gjrX+l*6-4K`Ni11g~'
    '6Z=(Ul1%oak<O8c`!Nb-cahkkvGUCzn&b?$9DG7+$suM~c`?+x2~{@=3EZ{#<Nqv7`Q?iheL#DCyX}M!#qRAP1=b<}%Zq`%Wah9{23rKxlN*7%m;~PN'
    '^w&P*O>@&S_v>jNJ$N9!sAV{$l2w;nKri;vj*31AtlC%{AHO_;Li6tnWi<u<TbthTmPG|>1geNLhpw~0diwXB3yJ>#(&jO#gJ}!3q>x$%$mR^>xUfWu'
    'Vnje`&!CVM-^L0q1~DIN1YF5Gf(=_D*MIy+dWT-EHYgwr*4ag`(CSqIJX(P=`y^1nFJ5fF7N0+V`r9MF6k}1_$WZv)a9KR$xUcx3J5Y2l)<wHQ;o}A9'
    'tX!exWuf5%ucdlIk2yX(XTG3(@7j{vbzAY+rHr`*MPTRxczW$Jw$ZINMlwGC!Eg@9xEJu@WS5{Qn&wGsc&!R_h@oc@Wru&Wx%DZkPu_yJiwkw^PU|Hw'
    '0bpjR4h(OKR-o70;EM$+i+viq?qw0bxe0D4hHSVH0t<839Y#02;5QL7*!M5toR^5M*xo+>rmwC+%pCrGF+O!hl|JgkaP=q*;6kSib7D&LK&j;h<J)Z^'
    'B`M&9$c9to@uR!^N0U!k`!(%m&9*-fP1*ZX+??=j?zUFC#)G+z0|^8M#bN|Ns(Bi8MYt&#8=`vzA(R99qx6vh$1a+T107T;AIe}1v11I_Up{_*{M)B*'
    'n7F?np$WaFAd@Zk{SFg?;b6RR|BJ3r$6kK><NyBU6L28`N2qd+=;T|K7idkzeXl8V(FSq8TJ-zl?N#?sfXul{DtiL4cU#J~j2eJhM-3Y)fXf53SoVuP'
    'pAUqpQyF0N^DvoqNyb#+_;ZsnuPIo!EvN}sed5J3b?HNT%&3g&reEI<r^71!dE+~;^DTB}78?Y4f4^oqOVtUb5D}&EVTwEI3EXdX*i-MyzR>WVzmX}u'
    'PiW#~eVZzX9e^>v?T`Pw>61s#G*#KUg{{eO<|<`TB{(ub+J1gmgf0Zgiz*E)Fiq>-GL;%<yE(VaO}9CH#1!D9Hz}xF-a=Gyf;UI|9Djh(#KK-w6m3%9'
    '9#k*iHM?$|lbjl}*+?r|_iCIxR!QBum#pQ$N7opSs~5k6`NBKMb@grP9R};dxrt1$E%o%ZCro0BOw{4wx<7zi00V)=MFZRN2$%eXvZxP0djX1Ri4R%%'
    '<+r#ECF8p2D{bnkya*SCj`Nr{*;cxAo@X(wgm@m+y1oguz68|?8E0Xoj9~R8u<Tqu4{D_i0yUR_lB4H5ru8$)s%lkr38od1;X#2TUEX8q@?~<aMCm0q'
    '&E11j)IF#o+ferehG;UQ6E4=2MptKX!oYZ>k?=M%JPdrbt%<6ZvoE=-CKlaYQ`(Ay4=u!BRA_coAktHWpwh9sA3<AV)2?B26mhoQfTh+0*kaY+a56nT'
    '#-zQ%gR0=RuPC_<c=Qoo(nA8=Xvgv@E(f{;A?<|KyFIi!lV_2PmeOrY#Wj%G6Y099N7I$CHhBo;cu&PO7eHjjmD{Z_G$BGDV)UandJF+mE<HpWkAj|^'
    'fI#sP!@Ai1<9~v-gix{`L%H^4$O&0@^5s#<+KUAQ$zwDot95PVUWtC_X#X)?9|_S_f`P7Yb%{#ceIf8~uwMe$Z<U}Kt14p!!ODvpgcBb4?LYooaDWTJ'
    '`Q%v9sFW;t!%_JF=a0(MSl8x$Sb7p|n_**QFL%OpWudWJ1H73>G^FjQ_aglZ5NG$8hqH<TxOHpX(ko8o$bo>wh{dXcBnu~-A4&<YosE@;w~bT;tZaon'
    '$k8>@FX9MgShNVidgTbAyb$5V?aMKWa&ZlSLNW4ZP%LK|)4S&m=$mKe;jYvez`bh-ATI-Wu}Ej(9x?<`7lJHP5ch5&h`tbHn}BE$Mp9SjLHK<dlgnrf'
    ';+ajw_kB!lU7+;w<^t*$*U9jWvkwX*tBVje#mk7?M8hvHAwQ-AaG5}=ti>}6;Gzx30FA#EEYn;7FgZ`gdZlR;<%J+{6vil38SZ%d+*z>A2F3s)T=7<4'
    '2Jm8+%%zfT05KPW98wS)dMsg6=)&%A2T-y5j+{W<;&kW>5XesvMceq3R6ldCtx@E~5FI6!@hp9L8U6A$259SW5S{@b5$DByGU=vN0YO}VFs2rPGq&|5'
    '-%n3rBLiz}+e;8qZq%42aQEgp1QcxdDMBs7HmFMwHqjn5O}IzpGi$&?@c83x70A^k<oLdgF<R+NII4MW2`t7#G=o+}XFy&=-X%-e>M}r>iy+c{<Yag>'
    '3t+=-3hd+92G2}UIzb)cV|R<HALYxf)@V2SF@k9wLlRsAX9Ba{thQGeEs7EY*3ur0qbhY30O}0Df*O1zM2)8*F*T)GuJ#y<{}j`Bi??pd+(4HDNOjBT'
    '_jMSU(piG#F#d$}t}xBfpee~z*sG*}->?r<mgdBTcq9K&rOL4EDz|&<e#0Fj1C1`Vr_m}y<5=~G%GgjqiTNkf(>km}p$Ur<(K^LYYuU`uj=0Sz*RHCR'
    '3M!R8Pwq`fFP6gwlaChb)yoLVUl^fGZtizozok_MssG0xM@vw*K?JGgcpB^epHF?5`ZW2oZk_Z2YhaqQ#G~NB+S$FNHd3_D5Jo4eZoQ--*={mVm`1f?'
    'YYh9*x=CJ`3GFCiu_-Hmx>{25R`(Q(Y1O24;7Aco?VUZ`^y_BbFHtcJ`%=Z+rqMy3uEorOV#mPkfIV$6#p0aC)0i^6=%W5pfTP`NOc2sJUgXaJCGUJ='
    'WHe3n=cbuT@D0ZS&~wWjh%#VnjL#j`C$w76?60f8Kx5K>YD^g=uQ5&^$GiO}MwP@aHa7Zzht>Y6VcGs)5ta^nD8-+{%IFPET}Df0{uEwLc5iAQ#!%UK'
    'VC(+9L*m{>j8AS^?Zsx%a?zE9j2B`;*Z3@Do5_>XVZQv7O8T#Hsfd({TVDA6bgx=dmO@lzD{Q?`NiFa=DwNFZE07>g$DzBr1tsCm1+&2A!ML&IE!9x6'
    ';4jd)S%C)s(qT&}j6e^D6}@D`aeeSeuE^^|He<gHi(%Vc_pm5WrY!&Z9R#r5ez<=tWQl>dFAI1?#5+lV?h6g}qbP-8+k?cKNK#H)0mUtaDGNXHLcD+l'
    's7mPuEwM!a3I&yN4jfM0MbF-GwcYj>;syJHUKf<AjXsf%%%9iq(bVO^K=_xX2+=ET!~S)&=Y`OSYPkXMdyHbcqTm${Yq9NwDet~`f0X9U!vPd>hc2)u'
    'Tv>AWBRZk1YIB|<2XYSw8#zWb7+`$;!y6>A7&OF^+A?~6CF88S4i3T^<4w^_f=@LC)|@`lkoxuV%>#z9Fjxy~<pr0D>yu&a12h|hc8Ol2Jxm;qrAm6Y'
    'G%beOU|O1cf18TUsZ1ZQ#Z9*#;DLLBasTz(cdVIS*C?qtAUt!WF*ZjGCa0K1IdfjO#0M3lDpuQ}IBt({->cH8ycR%!##34VDQr5+2mlGmrf5xy84#^q'
    'n@EEPp*%l|>;B`%2}%v)FSYY+G+cn3PWO9?wO)1~U~LqZ+{Q!*avgN`Lc(S@s(LFEw8{w^wu7NZ>-H2k-Va-<^6R2<cKg+)KY>;*+PC$S(r*oy0w{lf'
    '8zYOhaS1jsPG#nisp5wy3#J%D*jy0_SF)}ynX2o5D+(ljTY5a{H4iod#g}3^P@)@l8`-*yPatV*wHsCAqMLZTuGhn^FKAhH`PX0}Y0$V+!B%rOTVkNB'
    'Tp_=x@Pv?w(Rp<a1#wAi9hK0xwWp^TJ*2;n5rAVHsR7mqs?oJ_P*rI_wvsFg@hl`KRZ*kBeQ(ve0|3KG0NoO!V`{Bl#uIWV96HTwAx0>8h4&A|VL0Ja'
    'f=CD>J>z!}jrH`?K&lZ0(PE8UI~F?>Sxqpw$Q@(I@nnW2E4I;GWOgo%b?zXD8qWnYJ;HhrL^&CQ3de%G>3Q;-(*Z+r`8JPNwRu?-SicvP+UbyL#r%md'
    'x!j@0goD|C!1jEggQ5n%q@bqzq9*WwQ&n|Uy9V5FWB75k4A*s{Utkv@L%Az&3y|cg1l9<m)xs&zHMfzl?{oArBM5NS-#sPrL|xlH1`62<ISt;QRyP<='
    '&h^e7d{>y_1s-RR9Z(y9ci?3T#B85|v<K?2Ln&VQk5t}tvG#pt{6{>Hq8I>P1Aewyh<7)5!=;N*tX-`3i|%CO{vYg1@-{_kyaipg?Rlh!;Rcp_3e;fy'
    'hV>9(?giHyygDs`r55+93QPR1yTjKNy{RBB!6_btC{eWmq#1)?;te@Zudw#jm4(m=gbkFl5cL>h+hK0i5sxwvdvOS~$`;ERxz_418<t>}x2N6e*l$BP'
    'Nq>9*@p^}+?e|}b8f2={b7Sgv1Uura&(h}AZ%-^#+!wI*aJUa+0wZnyhOApjhAwMmv(f|!qo*wfaU435;4R)0ob$ha8-vwN<!s-Y+r{K%6H3n0D00HJ'
    'H8Bi^tZ5-D;XR}z<PUHXF!Gojlo+zsUG>{rxO)8&#iY_i5Wwmy>>J_UV4HP>8$qGN)c)zA0Pz&X?0$Jmh23Cloch(4wSk#neKC9QYwIaK^;>R~_nrm='
    'jO0oDhrim>*b3z%V>I`qqo|<WYQN^{dp~{|l}$oU;VyvUG~FpdFoD3&g^z3Ce$3OKkILZ!g>-Sua5VE2s95`-dzD&qg|X?+58Zm>dUyeWHlCnrvIOgX'
    '9VXzFW6iS1^yeeX4N$mCi4<j)A`V^y0w0E`*$Oiy7zfP(Nn3_Vv{{nt)AG=J0uxNG+&Fkj0BXiR2YV=xgpd?2@|T&pk8E*7)+xr{QCXtMYJkF(*Yg|K'
    '-V(#nHr)*-)}byL$m5`HAaxc^kTKjrnI$E|4n~qo+3N7{0!r|qG*$Hsian}MPX_AB1nTF;K=z<u&IIaO2Hq>5q2^PJ5;9C<GK=7e;WF&!knng1Wgm!x'
    'AEt3VJZFg~)|D2X0>_R^o346>eth&>XX8ZD@)=UnUp%EVNSC%^#W_m83q_GV0p5XwW023Dqxb&owkCQypIbM+T_3xdF*Hp$%ady(#&<1;AieJfW?`Bv'
    'LVWRLU3$A_Gi5r$mse?>;=sSGa=+x;6;2e)B3z3pw;_^H0>Kk08Hma(k-b5IM2BNv*I5$2Hic9!c?_sbIPcRJhN$R@MBk}YplGrb?q0-HK6TjE&YV(>'
    'ZiBSKC4a_b37ih^NUSo&+n!&6(L>-gCnAu|%x?NK?sO!^&u@P6ZeT(EK$2hH$-5z4Vlxrtryd_)Cn=RBS94-J*1n@Gb6ZT*={l(AY%qR|NsSG)<+!kN'
    'R)UEfV2(e6W=#{KGG>^`2SOF?2L{|#1t7cMYjmxbHfGA&`$alVl2g%+THJM1o!=Sww0Sabn=ikSVQ#&A(+%Fd*Mf@JfaWPn7d=p6UR7k81La}2!_t+a'
    '=r+<%+tO*{n03`~wI@=+`aRh&u!+;h&JwS%J}witVgA+J{3XoaxZk6JO=7bc|1xj{>7w5bONu>pVb-IlchHbQyf-f$jA*66b4}}SDhhC=_B4m%z{MF*'
    'LSLS`D{^|?byyUjkYWuw?UQWaFNdW4d7%Tn?XVGtO$wl5j+9pAvgn^c`5_{tZYb&(=>@6!B8nfkrH){wSD1)>l#YU<DM))}_8QrY9<~@4aD5%z$v3NE'
    'kC{|x(k&f?k-G6dY>qaws>vRRyA`UtPwHfUj+VvN$&sMFb7#>jTS9XKs_42vH{fc6A@28VMqvHjDk6m8I{lJql1l_5<qigi84;&DECp~pVe@qwJ*7iB'
    '^2!&#DOn;LEweKb<8oUH3HpKcgwqGTYW@~}b=AG+O7CUR0d35Vp)E14tYo(V#ikJL6aG~e%Xh$y56;;IlGD=B!q>>yBfXYnb4_5J$6&-pD|SL+8`+~F'
    'j;$f6lU>6JxLP?wW+(_TL?(I|{@AJ=FpaI4DABk;*%bYHfkjPxH3yCnez)Y|#a9K^<l#NW(x#ZMpK@!HqyHB$n=z*3HUo5)IZ@b_sas<L7Haymbm%pK'
    'D+~zA@_~Th!-uiYcN>k$uQ(^=N+PdRLSFf_8O_bg;Q*ev$(lXiqt93A=i6wZ>5lA9>%)qe?^Qp*I&4T@b1sYuQP91UphUJ<ilCtCNWP#FziDWEEipM~'
    'M5}~<-hidah+D6)f!&52X5>Hej~1_m7Hx_o!v?qU2J?nS9+B83gHS3Rl<@)5C07QwximI=+J$yUUNHyNB-cZ^ikuSb+3HXpiRvQp%l(e+2)sw|K8_Kz'
    '?Kg?5%Pm!0nX7$&gZZJr3v5Hh1zBVe#nhVpi#4dwFB8D3basSVlUVQao@_?htOGN}@yP_X2lPBncNkZ;4j`_4loU1h*Xytj2&Q>r<}v3$1K-+V7e~5?'
    'o3hdG*0*r?B)p1vq*O9N<0QYfQG6+SPQ&yy61}fxGhD6kmbjyl0%^bWmD0Jw?kg(xM4A-&NfW%{UdR*e1U!+doD9^zg@*NId(J4G(79?J#x{;XZG^Mk'
    '3a%tfmp>{~;|Ax&C^tMlY8Wb=Ai8u4L{u5Xkf7->nli;OI$&s@l3@j@Ll5E&<pTzU<jRCB2rOMGqZfthH`6+Mo^9`4U8W%1UY0ncq;us26Dq}z?k_LX'
    'y?+3THl=qq9$?Jm<#L=6i=d$E;kMiD9y}`{E=}yQwyjV-RK@QxGil|LUOjhw0nGAs;hf|4y?q(+sNdfbtZY+QmtHz<B)A-YSPTy?-Ef&Ss>c!25#z<L'
    'Rl>_Fe*&tC8|h(<ODxWXGi1fxwWa1fmbcp2gcih0pIu*f1M5(jXc-e$O1t_bCd9Q53gO1JXQ}xpR60e8N?}y<AR}%*fzaVWkQZXu;g}SQ4a>+>+FFcy'
    '6$Z2&^1V#-g-1>I%`~A5aQtMkH+mCT-py)5=fJW;yKhZ8%38}1SEKErwduSM1?MHj!6FnZ!W`=zsp9XV_D`btxL&<k-o4S~n=)D3x~aT-K!{K+1p_g#'
    'n*Tu26s|r=Lfd+;wXh1Pt;+~}YSb4jTqm4v>xBm$gyzZ#%^>$tf4C#bZ$e>-bXfLOW{yjvfFjnMFO7mu9!pr+x*r&R!~W4F<y!6g1?n&$7H^J+Q57}&'
    '62Nn(lrruWsWi|>N=eOLQ8KjkhYt_M`xW{|_QRUBnZm0xO#?bnZLHErS)kknb+_isIW!Vb7|WJxXknTG984#Z5&PKh4ECZRL_ztl_SqFSnPsu2h=U~+'
    '=Roaa!^2oCR!jKDosDCLdkN2@l0wy9nni1$j|m{tz=1pUQg`tcFgJRGfSb|Ey`X>sC+G;D+zL000xQt}-fe7ON;ZNVuuS8jXN?v~EG;H&z&Ql5l=fER'
    '#Q4T@_rmMpNGkqex4PmgQJ7eGL|))ER1ae7QWJ=C6oAk-U>hb7&}~0tzV%EBOW~y|>tV5%4&-reLw!1Q@kLwNaE@Jvr2cwg5S6J*(MaX<vUxZzvmUm$'
    'N7fGffF4A=f@(x8h0oRlVjUW^Y*zPh2)6Wu?}fD`0F}1}adIGTA5O4aE?1y+q6A{a0JLcQ_^*$jKmGXi^{4N@KK}Ih`q#%l{>Sg1K7WNjeEacVKiR#+'
    'kT{%%aD|$2BwTlzra`@wQFYsPm>^5##JQ4>P(X7%LV+!D2AkG4f3<Z2)*Y6Qa_0*@q#mMpf3(#GJPgQ%{@%DA4$)v-wFAP!?z4zhp)@&yfi?1hN-eid'
    'U>Mg=nx+A>^>jwM*w51Nea);eRjKPiXx?D)VnA3eZy!E7(bid#UA}(KU?s}}V|)4cfkEolGe{JtnLEn*3H}nsaYy&uI4;ShK8fCl3_bP8h`8&7z6$i^'
    'JpKIl>oEQ%N6)Q4&d@tA9996!28=KJ{SvM9AO8C7yCrJ;M^!ccXcrc(PG1x@-4a<@WyWfMe%GzB!jNA%N6wE5k&E0?bCD#!V1X=HjLBh|I}%M%*$~Kv'
    '<pLh1p8t$nSN|EQ;NDcFf3jp-8>*dyXSmyMI{Ru`8R*5ae-BLcHC2>DPQOD_07Z;G+`Ms|RtKY3Y37&ApHe@XaI85!D^E9FtZ<ymv&YF3grh)X<>rHn'
    ';U1rm?%@QN{9Um<VNbdO1NIO>1A_7hCb~g>sMLp!e&6kSuH%9=Y%nY~(1Fm}z?O1UH#?rS_!H_r#Uqh1b!>&V_$@Zr0Zvr<d?Jiv^kv!N0(C+EfbA<b'
    '8=!hq+P9`cTPHEqpg5LwJGQ^S;78%j<F~JG+(SWUVYATyWrf!&+yFAxUV(~5i-dN_Vu1(2+|yrvm8x)>jXyw0=WLhm1Yup{Clv^5cWL-jSoRAfoncAF'
    '9nf-z4uo|VPJ$i}<4w@)g&a9NqJA2McXrrS$H>v5aPAIMIYig%St`L%z;|d%#vr33h~20ChKefQpf!cW*$=}Vno?{zm={8tAAc-fMrP{TxqEfPiJKb`'
    '-m$LK36I+lUKTv-a5vxWKybJ^t%3WIZ^&s)b#Mw*uDUSS07E?nbOjECCoz4w+N0si*m4X_-TEr=oCQjyA)QYF<oVTd@I4%3DdJxr;ZLu>|NQwckH6tR'
    '{qobly>yUvozMVT=O4}G0D`20{OQx@$6tPX{QUZ_U%!1rc?8aygVHYAnpH`J>bF<@cq~Bv_s7?N`x)CudrIR6A!cMF1R9G*zkdDv`O~)vR286Rga9~7'
    'SI%?vm#;tn^7uNe6RgbZzkYlC1%xh(i2e;D2+><!xUI~7wL6hZ>4rLNkWg|sZ3`*uFQ5H}<IM+MrmBL^Rjbko@Kic%&3$&r#(@!hST>n101ol$25GH%'
    'tEr4e;n{b*Gi;5`mX(hqq$0{tD9y0cnxpl(R;XjA2)<*AVQqShfEE0Q<m13zeXtT&w!YdA=x!wc>7k%nuiMaC&{mc~HbH^;0melx3a;Zc0~1xJnPN1)'
    ';$=mGr7QZ%wNPB@A*bRv3`dNmRi!eGolb?V1br=%gL4I<FZ5sj3EK(v_&@y1KUr@SrY5$Ix4YOIOioaqp)4`6grm-y76ZrY2K6^r=1Fmi6gt+3MQpkQ'
    'Oz;LD0zBMkIRl*Hg~qyvkH3EV`t+8bf^AcSc-#KqfH5-#hT^)mP=<v`U2o2s2fPgh_27YPOCJis6){1zSlYM#Fr2U+;l8kqaTf!=rvnN~Ehx4w_N1Wz'
    'pvu=-B1Mdt^nx=(PCE>M1}=;b@%ltD(d-tp^v;tFH=P7&s|vDF2<{cCqEFiWxCuY5k{`1@x{H3T4g9}0f_p~UIC}T$8dv8zbhSSe?}rE6_z&dC0nNxh'
    'M6J;k8jht4Zbr9)b~V#oz3A2cFl@1h3Ops-epR&2_=*+XjCuns<4&c3wef%?{LcU`Kdc^dfIp;wwXabL;wX*q=ijhS;LYc+fBE{oSU!IHjyP2O(vo5%'
    '4My~8Yo($pJKUIgv+0YE!xk=b|M4S*&^r5yrK+*2kf&1RnCfb{xv?#eJ%P*TAOcQJN(qYlLh6*tdc$dOO?GgBrbSa1Trt}L6)i{f=As-~tl$y6=|6&S'
    '*mYw9S=tm~A`m$Orof-e);ruK^ATq+<-QSjAYai<;?`oYX*NZz#Wz}@ROMVaSYg##Unr|1atpGaj%glU<Zvmf;uQCtfQu|$siM*rpk%92RA2YO`b%??'
    'E&y;bCIT>`PK7sW7{mOy#`?TZg^t7_n1S>M*_fbikTi5U_$ZCe@%G{c2&hl^9}G@=X;FfvW5$X&d>AX;_pq+4=d_k0=s2dyP%J1ZWlU13#=Sm@yD^Ps'
    'YAyfKESXN0LAoKI&kjTE_gsykyW-Guz=6A|t$_+?0&}aXhz5Ibpo|)kC{M!1AX$0)i;mlXSoixCx(=!KvaA1v6{UB=mfM17Re4=Q8<yYcWun!&U6B~>'
    ';`Hw=5i+DuMnWjp^Yg@2)3`z-DGM$D6iHl_dL+>@Lqbn*E#8!se>Ykc^q9~gbaBFZ=-3YwtE(oq1KUfANy^-2$BJ<4Xh0>06X`W_4f>)h$L^YD-QlbE'
    'f@QWh&H?bo3U5{gGoXSAd;5KvJ<PS@su#!A2O7(6wG9;v!tqBF)$YzxSKg^ix)t&=YCV8WH%<$B@WPuE1->Dr3$BAjZ=NBH1Jhd#iggq!U;D@Yy1&6T'
    '3Bl@p)2-L8oZJ$nFg&HTMG>_<9~3{w&EJX7U<4C&P2{jP+nU2ZWc0Utkl(%(l)lSghJ_YPHMN0VCu~)N#&{|&Cu+2&zW~X6`N1*>16&ywt*u!NX3I7C'
    'M6N))#4rfk0l<4jQbJw@CaeSDVi+bgjc<wESnD{bhM*@LyARzP{4X|9a8u3=3aIu4Y9l3`TSVaR7N?JcOFl*=2y5o~vt7xjU<1*dh;pC)7<(>DycmZi'
    'dKNywidyQ3ae8}I2EY;({<F36OpSw7oEjQrk6SA8ik)DVTT<f{GMc@m2NBvxpbZr)ZKUW9OVm(PIW)nbAbz;*Hye0Fe}H9<&nb*<Gi<Ne5ap`Q24vx^'
    'lZ{cT{d34%4aEWDEj(HCKO6*%n+L^5)c!RP>^r!{7J=mT@Hcr0f-|#NxOdi;fXPEx6s<Bf*V#ix$R_J<y*O<skrYh{R2=dJT#+7FH9W%Ip<mm=YAqH0'
    'DZMa&KIxJo@9-X7Q@2QXUQQI8idjD&Zr-9uz%6J`2ef%q<v;%4V)=IQ_MQzfATlXeV5YU<p&T9JBRY223vYGsp~P*=6>qFoTzO*;@&KRlM^6pL{n@ZQ'
    'P_Vz)PF+!mB`3Y|TjR~Eu70U)$^u7N%1C_hhJ%u|y635@m>M2U$}g5x$cv$)a(S>$1mbg|D%m~@l9Q#d@{4wELKZ18Gp?fu@yO5ZL?IV%*wX*GW59ML'
    'wkvinV3QaLmw&>jOi*$uJr9!*D9Kg&+%#;NAf=C4<AP<_dwD2)Aq1)b8#}l;aD?ORDWC@-Lh@{9bS>ME7uvAUN}JMw0EE>?iU04&1=t;Kshb*lM?XGT'
    'SzuI`2LXLBM$J_!{xqOV2kSt9SMHH#8I;eFHRS~(9~z<V`YY}`f4D{uF0v6gUT?Y1K01gb;16iXw#TgN+Q!%9F>H$OZm=E~{nz^AuqGF{^FVhCX`ftL'
    'f3iDb69TjfqLl$X_0AJVJSK@>dLdA_h*y7ki-*)H`mDUWqF&ME99>$byt)a6Jh;E}>56<2AubeGnbYBY$ILgy?FWvqYe2eZhSjZ&(6#3-T#vyApW6V0'
    'Btpy*u5MQJ^wv)Ab7XMFw^-#DgEs7MKXf9>6W88TWW$5;2L-yvI~O7R))H<GZR2Ydm$|_u%j^W%Zj)XZdHIELR<`V=*(vXSQj!&GziYb{adyW~Ex|RW'
    '+xmD&GQ!$f9Ogpsuw1(R-hQ#(MJt`@zHL<;Yi(QXEZuT&0sGj@t@8B|q8GKH1Wfp4?@GLP`)$0jWu|>NmqO*39QUz1>56e+ZVp-EeYP_%%E21j`1c?8'
    'I;yzY&?^Bqe@UIIepo#edk}!_LRy*y^54d<Bg>dc`H5JHP~|=Rh3+M{b&F1wwOf*^ktS4;8P#s)kTJT<uu!(a01&s%NQ?Pa+`r2fPR>-!UZO*)FJ4s5'
    'Q)`1L7oDjZpH3rMbsj;NPgQ9NH_HX(ERb|z9TZ!9e-`Kkh^XA}@A=hhU6u=_#ki)G_JDhN0e`Q<@74V83vCIuq6`UMvW&I@w2u8SklhvdF4nE4GUabU'
    '@qmB1zrw@qrnS8i_s-6Mul#gJM>XRA@jnHoT$+<7mod3`hFpTI97t76DxV=$#rkeUtz&BS47DDSnwV5SJ7qH>u43X`0TyiCa9lkhWi?)_ZFlF-kP_1h'
    'O18(OMRH_Y;Y;~fCc8mgqy;xe*smASl>-170L1wbBin<QvMI3q3~7o1MFkZ39EvJlz=Mw<vA)SLCg7-Nao!q0(u6dck`5sHIS>sH)f6Hchj|VHN6}0n'
    'ssK^VLNEeoETlAzmvxBPgqG>GM-T7H=Q?vqA5R$gg%%eOdd9JjK3VGXl$6Wmzqt?JEZg!DO8?RD&9XHU?uG>ZE73=qnV&jXc#+gSM-Xc1>c+TDSoDYP'
    'hO`?M^)o3WN)u74=O_!~NJo3tq$9L8VQO0G=Cl~?Ry`r1-}*_#YgJ8t%QiWW%)rP}q|$RgcF;@Efof;BtGN2wjF%5o<*~DpiDe*`&k%bdZV7zUKH*C#'
    'z1>imiffN8J8U8Zmu7Lewa@2LW28vPtZe4lo(*%QrwJ~Z?X|%O#vwFHT!J96FpIOEGlWkIl}&jGga)x^G>FC}$2E1W=Mj8VW_B2j3^P%efoSA60b~*1'
    'dNX@>ouAn^H}vGut!ytQ_Xls+HZHza%U0CSE~P7oM*I2#8;iP`#MfhTm6AMo_;oaHIs0oj&MVyVCS+Y+bZX#jP|A-FAzb+Z*sl5+nJB<0e0T5B5ECIe'
    'nI4hMrAYKxJw>~J)i%C_09M#M55(_`y;1tDxfjjbIsnxHsJRqK7C1He)_53KDI8COx)jC~O5JPAhp$FQx|%(c)@YfW=&Sfdx8hQ2e^BC6IC%qny*&5h'
    '=n(3;A3JFN*FhO*M$PV*_IhATNdJ4_p^LrSma>|x1@-(P&&F}@`hs!z0Yh91%uDc$ii=8O76Hbt$J#ko_B&xIX-mg}3gmkwiR`M!SmbN?)NtJ``?Xs)'
    '85r}mDno@3F5G=a1$%+UW?dGM+;zRgKGDV(W~!<XXiuR-)&Y^LP8wMYE!fQ<v}*$=xCSMb!aQ0_1X%|p-0tGXub;nu`}Do|{`<Gze*VdtX&2Y0)$PvB'
    'hReN#ww~7*2uU4^?>IVk*BzE8>%_D9%x+TASymJE&lX&9=jldPuIIYc+7@FS$hM7^_Q9&Y{Nr0!OH2EN6}Cb+Uu6$N>20^R;suuKcv)B~2T<P!gpw<S'
    'Ek6{qFfM_(ctON3eUTpcN8O4DjIMqJMs@Vb7=U3y)W>fFPujoylWY2b0MS{%f0QwvxD+o~;isU+hh1vU6~bsDu0vyshOPTMYiYI|?Y4;q!8t`Axxmc>'
    '#!b4S$9SA@bsOQy*rK6BsP3?*lh4t*wid&|=0it~Wr;TU-fRK(JgqIJGNN`K3bY>>Iu`6KzQw3gkl&kki>fTdJNX~}S`^*&SdaUlwaXAuSOYp{;Ftu3'
    'zOe^eHek6~qMfwlyG-Hd{$W~!IO=7}nsYR`!!uRpaWR|g!B$LL;Gg!3weJ=pB&icFYsj8O7nk2HIfdv6JH`KQ5q7coK4Fzqs>b$jbN3BMBJtgvC2?GQ'
    'eHXt@_3ZmX#BPK92k+t+<5FVx>5VI}fd`n2$+O+wXJ>&rOYtY}_rYanj4c2&`n|UC4Ts_&nzDUoNgQ?SYe%tBZrASztXeoq2pgv(E{u66Fn8F2u&;uK'
    'EOSHX7OUmz3Qf)L4*d;g$_Vi!h4pR{fOE&&aDoJ&8c$g}-CZ=AlsI?|H=w~_jOJB;(^CdH_lBmB29!Y73ssj6pmJvr6`3+tWVmx0zB_N=y*ye1A$)y+'
    'hAT*0X9{I^pbmvju>Fc7MuQVYC_hKcMfF|LZISQM9Qg**?T6u70;bi<7wS`OQ%`AKZt*)xrg(=PNP8RXK-5*BZl=^oV>`Ajwj;r;nLWL1#!lq2?MN<@'
    '7ZSUpzm4BfJrT8HHXOcq9p7$sp2mM@8+}NcXNKw9j|fj4J}LF3)c)f!Z;@)A+!ChjWg82x6J?G+8h#sbA~o|{H?SdIm2zH9M2UX<nIZm8t1qGSA3Z8A'
    'TFfO8*iSuhiBj*NzdC*B3wS7(OG|WL`f`V+YDNhM-F`Xt<zeK<%056lOmp@ckbK5q?rVgn%F7NNVoq-4ch85-FS<Q>vd|6M6q`+P4-9yTAF*{b7dtF#'
    'SNWnlLOQY2N|`%)iJYyAeqDQcxok4OzNa7ndx3XYMb>e*x-CmffN~V3fbOtCBb8JKqI-N>mm5YnMF1kSTWRJ2BstkJnB?A%7pZ-vE8SM})R<sHfeSII'
    '3LkFgNt=5$9k;H05GiWh^7-9e&1vt1B?6YZ91HE$>RvWutY8AFtIM$_wt$@xF2WuQTSr@SNS5dY(!GqlMZ1lzy%neCtef1iE3`YOTjR<BH;NnavUWtl'
    'Q`=#+_C*k0)a6s{+;qE`kEnU&oUhlIlk8&Ns>;h~9WU9%wq5ZYacq#^4KV{#mr(kT#!Pea%_v>l{XF8xV*ck2eji(uIRte?L7J5V-!B&N^2z*<19bd!'
    '&hvYUS54pQef(AnafbLX(SY~Y@gg*r(fSYH#Oq<QQ?w;0JVCT3`V4{pM0}1)HKmAzfqi^(nlk@e4;<r8cc;G#@imLcT_Z{{^IHdejMrqkpsw7i!#>6<'
    'QqTX~!Jj^Uls=@d<ZebAv>rrFWAH~~HS=yH&dpy(+vaC{M($#{PPG>jy8++zqvJW5qLZntSSSKj_^7C!q4M8~KV{~KqOZgs`Tse4vn5G#BTMvGOy*%E'
    'bL3{GVK;j;H@EPxOy6kk9u<)?FV#654uz~?3m1`?`xpHueP8FX`^PR63P1td+>;{3Dli;J0mG`Q&_yoS5A;)G1U^{_H~<F_yN_ursYwDa@1a6l=?dj5'
    'qwYJIZH1R~$@1z!aBMmqMKRxB`nq%o`BG)1+;V0mwwhF6-nHwRDwx2G_*3!0KPlh>ZowRnlR2LGC(QBEH@LPbS+|Q(DGmB`fMcACR|n6C_CD(_PE;p5'
    '_qJgn>WoWv%pc8veWJ{oZ6hS*Nt2k}vjeva@?ZuxJxXVR$32<RMN7)Ik)jUyW_gWCymIdr%N92Kr{*izfPQIDr*DwU-Oe3YY%~6@gH?0eY`{1)vRN2g'
    'F|n}S-Fm10-gJW-b>A!I5+M|T^>n^A-}<XF`i3cHu}k1S#>;mm-wovCu@k;&Jmk22V`LSaoN4fj3!w?DaesX2AtL03P&}eUIYsdWlpqwFq6EA<&Y)yC'
    'b<qfZgpvy=MQzC`N?A#24y6etrYJp!`rWOe7+E?|VS0TURIoGE`9*5~@}c0L%MxOU*?&V2F$jC+mEGtk13QNNXV{GYOnCPL7a9{4zgWv<D+p*z2RK4I'
    't(%21J%WO~9id9hKuMHOaW2C{^bs<%4nGhP9NLs4EF6W@-8pOwr43(O6MW+m8&N*ZGRBhGc2V={s$g|H7{RPuO6dcuoG-V%_fhfDbe#YYcKSw@!Lc8b'
    'tS)3jOwA`p$yO}|Q*uhRTCV3xwanD9^=IW~XE@mq;GvDaH5(TgkSui3D-gW9LqrQtVSnq$2}2~S)5_?Zw*r2DXjy`4Fz!S2&T@Oh8Ga-z5Y}#k<0l%p'
    '3O_MdNPTj_7-3&;#hK?;t^iY%loqqf#)A;X?ASY<W;hBlqYbFWa~d?-?!`H!5ulM;<modaj8#>9;zPgBNhTW~RXoIT_DCtUCJ#g(c`CgabVyQ$`SzM4'
    'M78lR4W5SLF?1r+{#gX1P21BqaQr-+o87UChCQX=!R^OlO>p#p#m3ebn;nnAueG+S3Fre@c;~COdFn_R)8ayKFQ!;+0JC*x?;5fIhKhE}M44as#CC0i'
    'hBZ4dn?ClhXd%BFA|}R+=0+_*c#;LkLJmUIVUuca(`c9H`Q&W;!!Ws++z`JD)nZZQ(zZhtZF(R(7Wy7S*0DZ~liYq8ZXKAUVg$8-?GN2C?hL2y#Vy*w'
    'OCqL^!%D!6`8$7W7lw>!vTI&b{6Et7HN!Vw+b$|Ysess=FYOn6sq!IQzCmabFu{3u{{+nmP%_3u1LTSVCJ|7$6hGsqu=7jn-@bUr1krBK($EAOjDj)h'
    'm&IexTm()RL>64nw$52)8y4U8e59-b7?JRFeCmNAL*s?>aI!n^DC5!p6auOP|Ai?S<^rl?VDi_fA8Vg5qD^XX7kdrQ>@5c7_7QW!3V9MDs7!Uu6`z@|'
    'r=i<!-+Xmh;Uadjq9E7n;3NdO#&`X`kk#Rk+2<99t0&P6=YdG@gc<poTbiSBfDhp1!Ck3if^?3`vM~BiZlmzMkvV=EV*FrfS#1{u#fq`^-o5yAb<H6h'
    '@4Ms#;Nb-D$6<a1IKGE2AUU~!9%qbf%{QW?5=!t(dXue^d3Mog1?||B&-0$5=;w;7$$#>*j4m!xOI$M{X<hClYO96@oX3ew(#yHyD74}^DG!3g(m^)C'
    'INlQ(z=u+EN(n|@MQ}(DZ}Ec+G0E<S>FxY{I-PmUqlfQBE*}5$$({k-@B(^|FOcu-2thI6$pP+11lx5C>Eoz5TAKc+gL?p+wEDFb=%9?A6JqYn!*qV`'
    '$h4yb{<A)!kEii1xGdu(#ga?v46(bq;@b<dZ~yRs{(7aM(Ij=aAEurBb{kFy0#up+?w2WIFB11*k@)3+VEtBnZI1qo+e<XY!}zBq<r?jv<FJn8wy@Hn'
    '<>dA|=j?L4ZM!ttTD$id(a}c3J^8|F0I$@33o&B0Cuh6rJR`9-;(RoAw4W&__A~BGVAX&BFU1F^Hn95zAjD|zjgQ8TPSWMF^1vD21dyVUaXKMyQSwXk'
    '0(Lqe5kO1E=ND-A6&kc7z>8AblY51KyIni)uQhbQGjMxwIzWJu5sJ?hpF!UXfdCx=$(XpUXAcj%<c^j{z!TuvWyrJNbx0Kv0HjG!VjUP0$61^YBqHEF'
    'C=Kr&VlUdwaBLt}*#}oSjACViahDJ2jB~fS4V#<2w7C55|0UL}D7phDGcXL<z3GSR^R|KHi<UDw=31FO<u3=Yz6VYN2_QH@i+kwI;S=GU;OfD~fu$IK'
    'gCR2V9+{Yxu-(_9PVK`?Zi3Zao-hJBhSx~!KwzpY_G7YB><k{aq!B<4)?ZY3K!Pu&tMt&d7le7#FI02XoPOwbedzi-uu?`nnyjYc$<RUGnJC=suI&ll'
    'RUviGd-YjG^<#UmCvX(L89p`Tqz+A6JZTK1d^%d5ioL7%$UPpk&_T-4xa9E0Lj#nvvsZ%v%qS#=FkgTRcR(ctaz-N@h7JpgvJvxgAbTxQ>YA-<5`7mw'
    'oD)6bVH0H(Y3k8kjXelPqiWRlMg%@yLnb3PWOg;c^1^@+KD<dHVH{d|$_DP(ElwxJq2ASXH^lQopJTf$5A8=xOX`X2vo7nIE@Q|_5U3};J~hWg_*d~x'
    'bU#e5#KR?P%*-@)u}5AI=_bv+Vu~?dyDK$1Dp~=+9+giGN^i+7jG>|0qtGFKndQ>F3e>Y(ZhSMI>u3dSO<BfJ7@8Bb&pip-uZ6#IPGuS963mT}wS{QN'
    'mK$tbmgn)xJ24X@SHRLMA6t5<3fnedhiga8igB~!(X_uQ3M7Gnfipl1ZyD6xA(((m*&TRHm~R%NQyH|s!d2j0=pk%9vJv*9eWZ{}0_B|i!&Sm@qskwV'
    'miKAw(bnD~*X^JpX_fdf2Z7`a_V>94!z({DC>I(O)og<*%3X}w+u4b_^e&KL^56#%l)dQz5#UK4mMRkHUhV3=cyH~gejI^Gr}=!R;gfB>xsLO?BK5Ue'
    'f!DKeK_Y9??fLJvV?O<EQETeHSZ?3eo3szo{>k!KesiGP*t`-zWHA-&h7}?j({i-D9Lp83I?wK2M#unp68y%mbX88Z_D(p(dXHSIr(xe;693XV-urHE'
    'hiD}(SO9pX2xk;YKCbQQayyEs5%L-0;4HF#-!`33szpk;a~Xq<N}&~7NIZUa0cTyT+Wj;5sQ2#0JXN}$nISi9pZsL`Pzn)J=tw$q#ls;g7~Y1tOlA-5'
    ')8IINf3mpwdwL34vD$Npvf!@8QmsmchraXiBP>q5`&0uw!~nzyU@rQf=V4r>clo&roX6RRL34&{dakv7A6w8A#;2U>aiY<X(C;ctjpTJSOy{CbqJ@XJ'
    '!K4;rfBHi}NgZ@dR@f@ViY_fgiP<!};sBn?ul`@gyQo-dN=JpEMlJa1>}<$fpkzNYX`JM6!5N6mX_ttX_s^{m!SoJNUsW^1)$UKf9E)pb?<=pyJ$L7G'
    'q9dSV)cr7J8P5Cu0RJNhonP>G{MQtAq#GQZEaln_j)qe4{yq=N1xE;a5dYeB=r8WuAliNJgaBH;Dn?UiL%5_1^_ZSd<B!uG4@}3LuB<e2^_1WRI2sVJ'
    'rsIXJ0;Nm85`lH}H7M(KprS;-4jBFRYgw%UMIlmHJWtT?brB|uteYZsldX1B6)HPC0V_$USv`7nq0=U#J%Fgp8q5^N<+Djz)HEA>4O$gd+9J(9TZPp{'
    'ox&=biPetsf|iWo@nS5s*4fEEJ=&d#+lPbYHvP#yJh8aC6}>&${b1#UtGy0YVR?NQgDj}6ok<pn3d97Ut=4oqK;Yph-I^Kc))~mG-o4@N8EMqBG-Ax?'
    'RZ5r6ei$5v^aaES#UH0%VSjo_1SFGUUS>|aOFwi+a7~D32N_X>-)2y!W8ac$lE&8kFjN2BZ<EO|#0BFokpE@^x|Z>3ROQ0!;U2s!PWu}$7*|jk?_ChI'
    'OsR7%HPJCYMJELBvNR0QsDGA5afGapfl@`3s-K~hs$Z{+5f%M2R5Cr)l_f(Xe})Wc6R$k!iuxHEnFM_0oXC_vMTRt@SFVVN_!%OkfxB`V6IT8#jnpl$'
    'y$bTpt$%9OWO>Uion$BQpP`c}N3XJ^7%QsgGucxjhsWi+{fCuBaX9UP_aG96NQhMunDT|skf?}+TqEIK94EcdoFxVVZv7A4!CoBB|7$qKfOn~Man&OF'
    '`6^@X*dv4)C_+?@uw~reNvX=ycf@eYNALqD{R^w2?$VyVdR^#2Tgnf=c$NH%RH!*c-xW&G6nu8vgUs=SIe|RJ-o=6%ln*s{NHjk$dlR#G375EYor2eq'
    'f7mB9-XN>ZFEorS8Ad1YVueQr@OyVmuSw#43-`<9V%aKh41-C?96N(*+HryhRDoXXC1ttjZ!X@pUrazFgjog}^&*X69hcGg8xe`wpYhb2C*q?Tcp7;}'
    'SXCx8FpRfDhb(&(N`}5@-`tw^Mttj5DU>?AZEgKe*w&(I1j)5wi);TLgRu-KAvtjPkUir4Pk(@v@zjx)OG5H_BoC~*t*xoHTBxmK?Y0Vm86><u!b!IM'
    '<FEhn=f4y`zWn(0Uw<wBPl1VEeEluB3Po!kQ&Cg~Pu{D8c!<Nk?Uc_*&P=dv4nd1MyOr!j6j$M^g0JUGPtX4~W_xcibXoVG{}#hpNMLV^H|H3dfzRXy'
    'P)^b49wW^dyYHV4q^n>mdJ@4;wwJc@K{$V~GFT=dNJfw_b0Fz*V8NlCfE2Z)=RgWfg2kvMO)brWbiws2vwZ)nkO|CCTeBTBkVcA$yn_{isyR>vAW^+D'
    'lYnXh)pMX49;)mRdY|F#6Y@70-Gi@1_uwJAhs8+3h&7Cin0*-&o(R|$fO3i+2%Or5c=CG)<uNo)h(HyZqHvbmz2q@xa2*0~XWJ@>dVc+5PEa(e1X(c}'
    'v~Bm~0|>XhdbB$bOByMh4;9mq`^NMfcuF4nS4`CFjyo4&Bxd6gaLUrDGBmr>k<n%`vpzFetpIbo%ynjafrEg;kD`TNb`Zb;84rVCP`|x^uVLW1LIA7@'
    'E-Er7NJGeK9T|=IJ64Kcq9(~{f`SE|i@*lo=OeQHavC7sdPz8MjOL(Ogsyogi%>;P>eQbS5}v^UW^0fEUm%h+t+J-Id6-1#f(%xZu;yh<dO+biLb&9D'
    ';d#+*6pt<#3==+dc6c|y@qsmi)2Djrq-&Mr(a%%39S_TW(_`e>0WU{jLYrdu-0j0-FQRwQt(aFAvk*capKiD~cFCvwjov1R0(~K!f_Xkk`0?m0bG&95'
    'b3StM&g*ymIox>N@1#;DI&@&{n(opaA@6M9&wx|KYrDHV#iIi=DR_GWyQg-i^Jze>)vGtavOiYj8c@zxuhQ!F$0JM9aV#_e;Z1y|1}n*b%+P|+dI2q#'
    '(V9F1*aBKFqLZe@JXpARuKv*RC`wqu>v<HUnn@8gb#h@)>zO`YRn4LStoBP39%3M1aR{`ARlmz<ouZjeOi0`2k~|~~{jSH%#@+YICm($RCoEzVS@};e'
    'ich!K9((fCh&pI!fsKY<J$3C31(!q9fA^AHZfE@j3Z57W+nIJ@R^+GK@vYC*KTjJ7w;!VvJ7MVwy+;C@#8^!B&IK@9!R^-*=z4si!`95GHT>DtY?*2g'
    'r+8)aG--7jH|q?B01y(>R(IaZPc2U<oFfJIbb6Np<tZOi8X_h6P+!RTmgwXbdwk7e58XaL7Eig&-V^+t2CyCtAfm8l|NXn$FIK4WLfZax1+QcKLvif~'
    'U+s*pL3qWvMPm>7*YHx(eglq@D=Xz3dUV1pJvRYJ_KmRB1K0o^SiEFhR__NE66moJ?8Imq13el8{26NQlwYK1JH{~aVh>@Ywhb9M#-t80$~vN(G^2dw'
    'Hk8W%g~9GH>Q;<@9m3NOgjIyCX0a~&!Uz}4!F*hM5kJo5kFuxKzWd$HvUma$hFwPSy>}ZBeu9^r)=r^_1U5Yw-1Mj7(tv{s=m(13!K0Y+S;1_7wQ<Z6'
    'E=~7uyC?cAQqq;{G9>$X4FPC(8)0}!=G3kmlfd%PtlIMiB3!V$&9mhOWc+p)W0$6ja52N!rvF3{^LCrt^C{+uRAdf&ydA({3=?;T?8NrqFAIlo*tSg6'
    'c^Fq@GF?IYFq*NpOmuLyTh&u&w38ZxnUpcEhcR`XeX`}MZ^u3h?FlMLpEP&Z3PW~h8YAW7RJW&JHm5i8aY36SkAj#hQ4X(0qYUQ)ra7XZWr69~wMIw^'
    '8q3U5d%ku9Es+L7)ypk~8GR2eZ6N6$r?sG$)<Uqe^Sm1>Wjql#x{MrI7Dm^5Inif!8-FIcl_c-3B+E=Y2%`ILcu5u_y9?*G2M9s}N5v_o3U?zOhWQNw'
    'gZPwP*P(-WbV*3X9aw--%ldB1rhxp+?P+VzFyH#tUQF~n>{P=c70s@5UVd>^r;*dLRlPmeF$kF@;?@A}3AFX`ni%pA+jYX*M);FM&Yc|H#j@8{068UR'
    'MrRM;V1JEox0nF+1*|<TNqzzi(K-qBy98yXsNdZj@HoMJKV1k|O@S_|`K7(t(aw>ZebQn*0eTD39Mf(#0b}q?$}lmeakuG4$lcHrl1-5I=gSBYy^96B'
    '%1t58965&^P^Q90_yl2pc6nRIpn*VQ2IQ<GbIdWt$O$I+F6W?g-|Y!CT^aEE2r^i1njq6jN`tnHkf*N%G&8_S_l8S4xz&7IGc)@F945S_MtANE8w|X^'
    '_-OBLK^{J7E9E#k6!H{F^=+4ss@F(`<z}gsOVsS(L~5F?QS-xi_@vkZt<2wD^EF~_aUMPi`JtGn6;_ZVCD%xW#dP?@`O;S&yB$$etJJ``0D87KV(0WD'
    '^%|`hK<RM)wAi(<5FI(QM$XxV*x{E8l;GWsNUg4n_(W%yhFI#Ey`>LCPpr`kv0DzGj9Mg@h8HlaOD(9a>@>`81K4l}ng%ns)?(}t%d&t=0%$1M9JT%w'
    'w>6s{iZSBtaI(_Mc6>{`*iR?;^vKU++l;VuHDMUfIq2No^hA!0f`xQY_F!=LU(jsn^2ZOXzS|$#D7&F`S0XCBqRA3V9Nzqyo@XlC?as}CMKdNY$7Dj#'
    '1;(v7?1Jw48#%mt?w(~C-3GPv38tqkVql&(NqFtuvx&KE0eeSMiQWcp#Zv=P*8${~@J`@}ku4>g&8L-lp*D^gA};3HGP_A{#0s<PX!tdm8htn#jA+OO'
    '8nORK1F?Ib&dbX%Voru}XY7d4+cOojZYh~h0GNcMy3Mj0R(^-%pRg?Ux3-%~+=a`(?UQcr4rO8*fdLFA{Ym|`@?ISb<?;XqdVOwpjbkX|WCB_dTF;}w'
    '^NeVu@Dr=c7>>jkXWa>zcELawHUluz_R~Xbio~g7cl05;{b#GW&kpy~=$nsw3a5jq*_{EwvA+Uqk7@p}=Ba7DbOCYYIeIz{pvyX2iOk!#?9Se!)OTf8'
    '_kmBMqmG=Jj1i1M=2G)A7d!_H(H|{}UT0w_<by%)&lGJ>l%w;Ay9TdJP?8?HElkA`1vyV?&jqK1YBJHIBoQ7_sp1m`df>s7L6536Dq|)KKtfhhP|L>g'
    'I&xh`(`<bIx3k`njC<@|Q)p<Or1JLo+aVD-`CZ9P6orF0@Zf{r`^lPBHo5D5G*^I2v>tp8=lQf5WU@i)`4%rTIEME8s97&bg+O`IGbI9qFW806gz=2g'
    '+hE_gXI5{zu^X%a#prbu1FTNE<fW`}r&=bWnnEd)%iF1Z4H*!VRxmN=bc;HR^l>9+Zx8D8AXe#@c-yMK`Kp#lNLyjE=W&dll(%!S{7cO2V|S2Q-9d9Z'
    '=fY$ehl3?5<_ox*JUs<%57`RtF0gGWkniXMySr`EtLV}VXW9tn?nbBzK=fTOD%)M*2&A!_Rpmv7bu@U_em8W~dUdz;x<Ck`Ps8=LOSZAPi`PX_1Q^n|'
    'wat+zh`SV|CEzd)ek)^WU-<ah9i1_PqXD>;F`O^p*)l%xpd$j`Irl0zd=P`;#pM;ZjoCBzJD$NY_zui|bGUuS&wt*Cl2h(a$6G$b{RNSy`+axK0w+9g'
    'v{2-c<IHGWZH)RAA=!IZztz!vUMk_P=id1`y0_<v?P6ky)$iLFGf@hsIE|m?_k_10c?+>?3&eOd@_4{3p~)oXoPFt!J_$KEqTkSM=TqV8+mnGQ#wnhg'
    '{p%aXq{Y_{XOwX1_Z#41hR(8EU03)AaeqNgd98OaatV|IRd2}XsD`-XCpYuX+zEI#I3q5>F*JHZTkt1`c9f2q$9quWAgvAW5xmF|d9l~7jxnyBENmqS'
    'yV|eO0otuKU%SE3TG30aj+ukWp(@zE+S(ht`0!!huH<M8>$$dX4E|3rn}A^LQj(+A(-VXT^P!j=jZnR~d}(j~<>i~KRY)ja8q8D_@=b)$VE#S+@Cod1'
    'Kcb{W)}-IIQgIR8|HLC!<bfUh)*pAp_I3%B6pyAdme2!s&VDjy=mZf<;H7nkPS6s~5UQy|4-j&it0@(pjL!pB>7IR=p(E+D9&lvX_XoSIX2_UGrh0&k'
    '7lo&E1f%XeVBPMH7ni<pj+P=?4-Anr_HqPAr1~qxQ*!4lgd7#&H%-6Q{x^J=`gT_FQJ4|{6!y=41G>an@BZldb~aR-CACect=^@2_kN4}QR;5DkRD6j'
    'jq>YU-tK*Ku4|(bFJJh8J+KrYXeB`oKz57avyy9XbbNZhpF78AD!^uE=)Pkt(Z!3<nM?6Y^H%(~%p_yq*}pyAs5T$f#trb_u=3OH9Ah|C3LnISH+J}5'
    '2#ISs3@@3TA%Xd&IXZDQ2Fl&{(#V*(p~5IaG5xytIPEw6_jvhu*x}qa=l<N#`!w{1{*D4-IyAzb(U|gme>lUr%1C;slWC$^vZAl$E4w_Hg!lhW=j@P%'
    'Ml+66$tQReA8mh@@mbCwV+o-MGh05iGPnB+>QaJ&{kh?F^gA(%&vu-~8vE={Kv`8_tY1b#2VM)JBp#y_Ucsbg7?N9be?B&XCFJBo<f0D3c8+D_20>m>'
    'XdZr_8p?^5l*~z;c+H5GdXSb~fc7IzM(TyvoG7t}D8;DUfy!XTs_5#W+?_F{GJ7~&8q3t;W>F=_glKR8TZRnGC@sSWOxbIOd12)(zO*}S`MyB~l4EWm'
    'Cin$i;V62*q2Z;_6-lx#5tFUnrW%K!YjNf;WX8+*tQjvw0!xi}-CrqeJ`+ESeN6(@#h3_GAD>NPPdC5>NByM1N9pFcvn=T64dkQ(GhP=M0T`@+;bm~w'
    'hmh@T4O2$r*HI84=@UvdR|*_H3dM5RNPJSlbHVA0tbEgY>(a~V_R<~u0RkLYS>c5D07O9t-wRd+8s726>p4VBWQCctT++sU_tk~4bmrLEkhVS<n^RfU'
    '+<$9N0nl-2T|#Y4iQ$F3iUOvu$4Sp5JxfKAh7iN*S{A8uRw`gMo)EJ4;8SmD`;<0E494lnuqp-^5v8uG$D-%7<o|SN&$yg0vXG=j9|WHyzEE6KTRR4V'
    'mXums9ubm!v0-`*A|k{j5VHo60ivZKr-OP8;tj73TUzVTj96Xk0iuHQxjoLrDnthigfJn;COH@0*6}vC!AIISbWfr*EE)}A1dYRA{`3F)OL6`EpMNc$'
    '{``->6s^x-dPy^w9=fxr44Usb6Gr^3)9waVu~x9{E<O${EKUdq8_&V+TuFRCa6l;|uzS4@-Jy#FAiU_NgK(ND`ch#EA>S@vgxmXK+b8i`4&4Cp&688L'
    'oIe?{9%7y(H&IyATKlC4>`Kf#n<i1hkO2M)h9?BD!;jt$hh5E0qC0fs-P?m8zCq->d}({{>Y9^T8?M|HGULM{CzgjP6X~J47@zF0e>tEgiMi-~xDlKc'
    '_66TFiWdk$WJmL~jS*cB2W~1uX7ME^BLvlZ;LV&GKPW@u%kAM10=3$;<3je>A%+ON%^L(!bru}iL&4$5nK4n*Dix0{VU+1EW`y&pq&y6QJ{yAhop*Ko'
    'nY;L~-dPA$QS=<=%lhkIfB*Mii}Rm<{oCLF_`??d@z+26T5SIO>yJNqjN(-iKB!FNn{ObL9>hz7^vYHYJcAxcdEBcHZ0t_Z(sRO%vqZ;%YGQO6F!xMO'
    ')IdNJ-VH=^g(Dc)8a=6<2a;3ia)1bO^8w_JW6@7PEc!)$bw>C*ZlcE2NYv2d#Di0nTV-gb`gMcV`|hP<-1BB!2&b=ce@!g((YXu5enNsyyg-fzJ2r;G'
    '=X(|4_n6@_cd|YnUSo(i%PQ4#)hU$QnZ+E)E~pxxf78D7_QZ4c2i>I`UMx>($FDBaxGQI-wq)y0R`X01czWdyYXL)+_EIim+?!#53}jR92E*IbPe3uC'
    'z?oIlHI<~-X4P8Xs%Pq(SphlYe-t*YoJ?-n=vRSE-BZ7VU13-E@pAqQCqnnNhZGGLFm?k*-XR?gRgwdPOg_*q?7ZQVe@HolhhIgKf^!|dSqxoI=f<*2'
    '#N=a&*z?J8MVF4u9HJubEmfJH!6sLeh*l?K2s!qi8ql2{L5cnXJ1>zNGTZob({^P^ih6_N;}qHUrz>cqgC0OabWtrlZZBk6Q%ryNg?4e`acTJu%PF3U'
    'SHBYNzO~Z)g|xDmMH~Mb&V7Fk1jM;JvLHALCLL$v;-%ql6d2Jt96(h9^282?s}T@Y6Od@W0iOit8Fc|$ppOX!T~gEDPh>)2@62h^cC{ds<Hwgf3{?qO'
    'CC%}2xw$H9uI34ftu_Z$9dJ{!sKB*jBVw#*pW*ua8G;sjCIo6%k$cc=wzd#SZA##Rcf+){j<s1;OGX<2m(YLy+YP*=#$hj$*W2?(P%`JzRNP00b~o=O'
    '#Oro4wq*;?9a%MB61*Wi%kbV>zKiJeP4~R1?MSd_Rp8E+L`w82Y9!vWmuNU+FB0v%sCb+Q$QltT8B8)U(L*^I(`liE3*qhe(`E-?MX^!E9N4S=0L`Lc'
    'e1^FYk+H-OThm#Bn*=<LZwW{cNZ(lqhsvG_73mfA-FOnwOBwhQ&u^YL$C6XFRn+Is$>jK`>ZiA5WOW2e7hHz7Xet%h%~YdlK_c3%3(N_TG?XNlvDXAX'
    'cfzab^T}C+@VJ><Sks!A!g6*^ha){*)UD<Cf}?~K7}`F&sgkA}h@K7%5f7^7a|_|>H2IJayxBiDU+w68!L`|iOYd{r5LGTK6$n|`rKSb0)<=RX9eWT3'
    'af_?W5)L&(+b&QzGCQ#i(U3F-e6{Thm|wa4wEZ^XLv9|IJh_~CySw<uJKE+h70tyZJwnmo#d&y!rQeEL1gB^Twjv9ivj4C9ntT{Gq)YC}NGnMS`Zl#N'
    'zsU6-nO1cXz6#D#3MeJS!Y&}09osfOqJt$MWE~&4q%POS)iFV>^YjCj#t|~w6VjR@B%?PRT>973h)~ShM1f$>ZB=SII9=F~x$pnl4E8&KiEtX$2o-S#'
    'BwK)9AYD@N>_iF9y-)kOtxFNqJ+}L5X=_9Dkki4%%d~%LoK241vHbuW7AO$4D(*A6w9KFFS+;dUIat9H>_!3RNeG$TAnS{~fRM0O*afq+u+g@xO7vHP'
    '4sg6Je8`2U%s96@2*DDWVB8frZLu}WtGB;}d2_Q?d=RSm$LHeLfBoxU-aTv7o{0n33Crm+%z!e*9<qt&Z2%GXtkpayGS2Op<PUA;D<`r?dog8rCbe^m'
    'pVry+b9J_fSID~IM#WdpVwWSwbPX>BF;4lpeGBF{KI4F+YP;d`!vviGoWyU<=7hrVd22a?;3^99N+&Sbh<6tV<FR+vvuM%DeR2$MeTd?r@gsx=A1-U`'
    ')e8LRU6!j!z$yivucYtv6O$^QQwmK#K%PjGAnk8{AfBXAs5^Enkd}+$LnehEZvXjj;I?`_Ij-`XC#HPLPmK8=0wvue=A(jt?J(fsn<v(gWLPD#RubT)'
    'If4_yrNNXcBp*ICw05fUE;IxR@nFj~(Bf9*J=gFuIGGm!B^tEW6?=pRxVq;Qh8+o;^&=$6Lp>v;&ihEnJml&gL#G5!cS}Z<i9$fE$iB(0d}_TLhGnWs'
    'E*8eY=VbQ&P3dDe5I6^XrXQH%6Mv2vWDxM=T2(Ul913<z_8@&=VWK~`Jr;ilDwZfJKee0qQFxWz{iqNt6R>;wp!xx9D(qYIDWZ*d1bu4bT(lka+B|?w'
    'qhSJ^c7<8m6~V(ekGdcuiR1#P?~uUI4+Krtfm-W);HmRv(58EU-Dw)_;M2yFt%&F83ObDw5h-kG-n4#T?6c#nMSBoL^9RsL)Ch|XT6Z3_?sL6&mJ)LU'
    'k;Nwmu~>uWolZh^RB{bAK7WtN?kXSn+SF?X;IzzD^$0wgJv;C&W;bA{eSQ4A$S9p@pKJ9%H^!qdL^ywB%1Nbap)EBnAH+KDXdgHGkZ6n-*6TRV8FV?I'
    'chrIU(GI|ti3I4Z;p|TH`B_7K=;@?<XaHSE>;OK}=T~>y$0wn*W!mdXJoNQnY1W|}m!JNr<(JV(j?2|0FpkR$*hGt8-uZa+pr&uXCMVS+n-2~5zsIHf'
    'lFZs8H)|Th*TIwTUcKX@-)8pdF9&M=k)tOwosUG4vtB$*1nu`D4f+ng+7Se2U(qG!JxjJ<2%(nn!4X067t`j@O=J<bQ_)RQ_&g`679fN|I(q3cv-Lmk'
    'Di-2&L&9!-7mUfalKdCph5Y=?MSNJrxZ!9NYX{~TQc^WTus3%$SwX4z44w^~MEH@R59$v{8*uGj;r(<=+JGbD1&|0l(sP{>G-LJ362Vg<40?*5Uc<`7'
    '$l5LFOLr}hrc>Kdh<Hp3@y>(V-~Mz7lI8OzVct@%bChw<u0WSAB|TD~H!>1Q*`V?+6=H-E8g#F4vq&Kk)coc0yy1~T+6%TwBRipdFGQe`Trc{OUgQH3'
    'yF#}FuO1<TS4BdASS0fhfhd?IH0ZSE^&>QhQ+RGAHz7i{5>k&;V&~p2)#J$}%5*Go+w@K!_wJqr7-x0xcRWAugbHeFp_L?U7B2f#_u;+v_PyPj#jd$+'
    '`{ILp`2L$$FFhwyoQgT?9~EZVk4$ej582K5V=Z^zzy~#_2lg1AV6<mFX4S^oS=NiXivF6I-E!VjBzd?W@jSh%$zxlW<pT<957aef4^h<$FljvH)O*8t'
    'JY+b&c+7Twm$2Erk&+D5%9VlI-z<OLdGmM=t{5dzR;b;d2cy<2pi~u|Dk?7o&h<0=KfElrK@jTeO2mO;<ucy&I6RMgY-@+>*ikAFCeab%HFzFzdeup4'
    'BiP<o-u+Wb(l<gvr)Ph~p&3-jph+-W;Tz$~72KbSPxg|7{PoxH!F$!5Q9lfkx=a6Xj$e#6XKJi9aX#$-NZ|>dKZqv+o3wc~1CQ?mJOd>txsHl3d(to^'
    ';;O2JoPCfDW*#w_GbFEeSEu0g)B`ZCg&|e>l6oKM6>jz5TfwCOWHdu_AjM=uwkLAo*#;sQweG0^BAMES2nL}{=H1#bmV6{%`N$%eHh6<Cx2FIOEQm8U'
    'PAG3A#p{A>Yy4+bxBF)Mwz=&*xg#r^11pn>WYg)lUl^8;TljG<9XeS$imhQS&of3QP<{`ZVb58gxm*R~lAm@x-E^vafW-I<vTY~p2PtGEzULoO&C>XO'
    'o9X@cwD)oeH}3NHY|zCTEX8g1`68M9R^0BBo^JMOPxUf#)K|bIK(!7`k!_dvcDBu+f1LEU?b7~Mt5`k;CP;G!nL_j7<+F<s{OBf~O~+O*!bTm<2ew{A'
    'rLfUKF*`EoCNSw{dP-=v3><YdJ0dh+$ECny-ql$oGWnCP9@DPM6(Ap>5|~(prMMyQ=xlq=w`$VQA?>GL#_<s-L8&!V2Fu5ji+Lb~1ng#Jf5^tvtb)fK'
    '&Ih)-i%r3otk>o2=U$rqA>>5QuxXd;H6$Ow2xrQ5Tnc>2M|gI{Ca!TAH0i35D?mO%B{01T%WzG;WX&=XtVO52ZPM=Yb@aH;P1b34ADp54yyvs4IOrRv'
    '{qEC#n^i0y0~2I@2bn?hBbTpyu-9!)yPY1rM%^Va?tFFx8}SgH2uQ6}l0>OomgN%PngCW~HqSgH*kl8sHu)K_DyHPfHgCPuGtm4IbB07JSeZu%@DV)0'
    '>&Ng3ft;-<=LAc<=tuodxe8G)qj-)<K-C&3L5+^+Inj0ggjVSRE!i55=coi$uEP@C=+xtLbCI4rqt0fBo)q`Md<0Ff{602-56!m2c-oR4=v-0^{3T7>'
    'Qh{ou47gnkF@fUwZGVCKE({UHfrN^@q^fFKm2BD$ZSp#u=yDG|q<20zORW&2&g+s7!8YcQXvHhFj!iYd34!E|+hhbNFkVpE9$FDmWQ<VN>+PZ~0=0N7'
    'rKX4pj6I@iOx~prAgr-u-qV8%9qC9`V{yg-l`fa;#2v%ZrDFejN}E%&II#7|TRUZ~v1Qb#v2)s(o`vCv=FMdXVp*->Mhy;WgJu?poNv}%#xIxk3T)im'
    'E^V(`00v$7sL@^4s<{WwnwF>hpqiOT>1RZ$1eOaTWjP1)Uz+1~Igrf*sk0+AdXVF#aTgoxqwN7L6CkgbKQcnjx)z$0+3X?xdtxJZ7_v6S+6aY4!6Awp'
    'iVKl>%)oAUiK!E(0sx5U+NyLp1)Qb;I%iNu4uEshUw|b`0jFyxzTx>^#ff%~D43B=6yJI<R%NRE?YH6`LFZD5iafjcx@TALtvCyAVnnjKD(_lehv(G`'
    '7+)JtMApGep%R!T3GXn}-ut0{O>!EX&o#MBfn!58t9%VgT#pIS*mb4x`_Q5{rh&S4;b=ZTDNX@@E5}5v8h;$X45RHgeg+0|4Ui}dW1^~V*EsLLiEA|='
    'SPZ|VSsDDzd`yy|JD<R}xOb5|Fvp!;(w5oU95?oGVBwwGs|+yu2W5#IAVBI7gH^jX!;nWP<SMp6EyxAQR4Z;8jwd<qYajO%Bd(DGduX)|wg{Cnc$<4m'
    '?HGyF#i?{L5JHAhu=#0ES8sAW*kSGxUw=01>yOngO&G+(hO8PzqBjo@(Dqi}G`DMiIV66e7L{bO66>hsv>0w38lRRXGCb~YcKHp1kGyxjNTbV&G9R;_'
    'HpIU6?eo;~^`c~W=QNyQm^Hgn4tOXy;|=F!U0k~TOH;%UI00U3lZ1CHM~<$Uy(zNG-Y~sp-i-*Hw^IKM__t@QkN@%SOhDSiqB#$z*U!bTfBpMsXukct'
    '`TXjM_E(JNj=4!$beOQccE3@6w^uIzk${fP3#A*Zrnr#jDxqK{)<|MtwXWjxig$`iK0UX3zDX;2#NaU)`}*v&&)w{}Uizc+tB3ciuQynqS*)S+&E80c'
    'LM4|Te)0H=H>;UOj2O+=cgZm9A^ICtKWC1Hsi!?5LD6tTnVg*7aVPrF!RwjN3OkJA1F*;A17HNIncx5eXU3o5_@c_=R8p`a@UQ$-d;lV7k1#-9KTUB`'
    ')IX8eFB<0!jZ?~8XgS4P^Wi7*g#^2+rug8GE6GaVp)HG{%OKsFd?4K#=IBBAlS^M5&uqoUXEp|v?;Ff>3~<<3S)iZctW}vMCQZ|+#$=>vpc7avxHl~H'
    'u#c{&!@$+6pdf>eAs)Kpwuca-jXl?n_8hW9Xosc1st$IwlSwge&?9Xn)KFCEd>F12ZRm2e7s-WEq*H(w{3Sd4KR~c?7xC%4<1F5a>um$2kDk`vvL?(2'
    '?-Rla1eepW>45n_MZH_H_CCc=znrB>5{7a290fLB;go{|We>Blw~b9WFi1byKEHkU2lVB((;>{e0{>(et@pOOLF_K)`U96=c350qm^3jeiF?toly?I}'
    '$o$<693W8(@ZAnck7x50B5+YEy-*y5viRl$FJNsI!gGeA6&3BOwjcLP_vXWcy|X|4A%U2Xn>7?i)`)S0!1&_B{#^2pgk~}<j8M!P9PF%(8x%1=Stzct'
    'BKA_YTMPwW1<d`Y$ksS&GCrq=Vn}sPx9qbEHVnxl+9QOnx!V0ycw1T|QUyn{P=uvHX$m0epY0YC7MCpqhgLeo5DUc|{q@%E-134@X5c8uCNv<MtDm?~'
    'T)K$1CpdQCuOHl<WDnht9nw+aIxLt_e1SMTmU$TLx;*;KMjL08VFw%{q5Sw!y!%j9K!PhCuC-7MV@e-IImKF+XK(K!xwLRB?cIKPE{!a%lBgOJSfx;0'
    'j+vh>7+X8081%)J6oX?VHOZ=b83r?Kttp^47X6jnUE13=W?+CxIUOmG24YR(FY1b;IdIJNWDbfLa3D~TL`4QFtb)q%kR>WnSr^b`D(8v`+C^o`k)xK$'
    '=*$%}EL{4lU_)S&gY&nBRU#D;spg0j#pQD|2c&?9u9yI0JfnGJ#Ysfs=_Q-yn{T`p4U<W5{#p50*ntR*T&3^ISv*l1B7s^{JY83OoD=jbbK|5*iqIA4'
    'Ql<<2bN)8x<`?JNd)iNEQlv9ie4M5|n(gizlcVU0uJ{vO?9R=YS`y98l>f{?wzrsEGLWPg|Azk*$&fiAqee4rer~3x{t#DhZb`Y*OUm!-PQ9)3p&q*K'
    'oB+6t;RUehSh*))0j<pNwIQ2iShg}nnrOC%#V5-?b&j#WxTtCFO~VCk2Dwe7&??%8V$sKo_d(}_gP)5>55b`zV;9*Hx9yAlIGF6uA2@rroOAW#w~q;V'
    '8S}II;%FZSbL&&G@2s%FV*UgyPE}4M5pPc5lnv2T7}F^cLn=yFd}@H(hfEbo^7L?!Kn#GxXFEtv!qF5-%oD39Y;GkU=aa@flJFvBQcPw!PH=DK04vqp'
    'ntfuajd-o&Tw`&C`*&B{hYd0QS8v)`kyu{cRV8jL2kwG`xo{ZFWRxhrSUIw!e{R~&M^MEvPd;kYpTxYAk$LC#bN6}w84Wv|0m2kx&MD|dH6bC0gjh9$'
    '4cev<JkeQ*mO3L=5wU9B3YK=rm%w;I|5#;4E2w5PRIFYl`Z)>E0&7%y2FJ~SJ&wneImNf&<-DQorTBV!b160KT@Ctx49^J9*2eDiVwdN(84X=di{^o^'
    '&cjod&%=`vF0sYr`xMbscSbG<V;GxQRB}R$=9bgk+GD9Xx>IC(+Cz4wBj*{4CXd&n6<VCXI3^a8p0xeY-arMBg%%yHR!K7^%PBdE%Mwhsow;^B85u*P'
    'ZN%!x_zeWzK@HvjDVPhcd6^NEL{P4;{M#|)VlTLwjx(LUSj;|=mStoA3yEaibg{_?Ry6h1`qURs=$PxC<OjRTxm0XIjwZjV)>db@7F@(Bm#zrbDr$Ib'
    'Pa#-sl@g}t&sg;NaANe4p|Q7JlOf2`-W6vN1Yl(h-0q9xsKt;2FspB<zWGU~+JrX(cpc)KsMNdsLA5%pv>tPO-BPNIzQ@O2x0eL0(QrVcNrt;tSr*(A'
    '39z`6eG%i$(bW#nTo4nCvp6c8nT312%APF`JU)YuUSN=<2=ZAdc9-th_K6Pe=+su}3BUveOT^+BL`TOM$O*_-fp(UFH8@Qu2%;e7D1cU@ouCY%%nS;Y'
    '=}sx)n{CGUn(*}u-n~|~lW~$XPLj>ShM~Dw4)8enPDvvr=SJ$_GJB6fl0b3><TQR`j3SJh!3@v0;{;O?rkcTE$ovr~2D0+SG3fbrn2ds=0aSAXc(yxu'
    'd)v2@g{Elu)ZFmBG_dcJ2_F%V)G`)P4c>P96Py~B;kAc`u^ao{h3*B%AF#hD{*lnobcpG(KT`A|0u3L$Ky~|@73*A3LI^%5>x@uEgsRDnn+G3@A<(Y<'
    'H!ClELwf1A>F9pwi%Gv+#Q8~ta(fEGX<~1sT4jV75n|&DE}$>Zr|Xx^9^98)G(5uBBCc2O{*}c}by;yJMQLOwnoZvUOU@2O@uB<MzyAKmAAkLG@#7!='
    '`sZJZzyI}*KmPLP{E*`;7!HKu8PsQ|;tQx5kM<%cc;#Kw5fJm!h2r?SvIJ6@&&8~ZfVrsrr<d+)lL0gV^jsMKW`78e$N^#qVrC%wF1>SkO!pOv<NB8d'
    'VyGnu5!ab7P9BHed|tP$<x|mzj;rDWSMTm0D?Jy#z=q1DMjjK~gkmnHJId#mHfQz@P0Kj^yBSW$%NHj+ec&$boBEg$mqZ*Fk-Y(~As_)}c=1LY-jDe$'
    'iSizZ)I_9?@z<tLy}o)BT6__;BU=jDU_}<oMEKBX0YsyK(gn4VzNClR^jYZEMzvG(=un$xj*sI}$6!{WE`2Mn$FZx)N!r&I)N3I|wR2v}R!EEmlGHa4'
    'y)Nec2vw;Q-mrLi)=&+Zq9u*Lx7}StL%M|im~}H=4!-!)`aJtlTL_asF}Mf~C~810Df|PG3`hg0^b*jIN7kaIPg)8%<#us6L@!Bdu9QwqJ$J>WC%gnB'
    '>9E8injHCKJe)Qupo#~X^hUTXn(fC=Y)_|imv&6i5u%uw4YNHU%gcGXjbdQMP+YxYS_o_mH_J97Q4xvi9*MyJw<;x3QSU1evsBnQ^s!34)trCCC!}B@'
    'pFk-3l-AGPYqN8)8C=}xlMBu4_Sx6{J1Z==RJ{E9*FXOHx8lcNe*2d{{_kIZ{P8dU`}cn%%8G|{d!g9c3)>2Ac5!T{QM})imq7_8N-+=&=0A-=PI7L<'
    'hj$L6_0A#62|FgFjEXU3u$ac6c=|t2A?+^6Q7lVGf%uwHgkM*jytO3e3~1TyaxWZu3?!k852AXPCe{2}QZYHDFD~?om*X6D0Z+Y3v_8|XQiYhuYr^Xl'
    'JRhg94|qd(vy893Wx)hq5xiQ4`~OrB+h>MIO+@NtA_npkOwlnPr!Ov`6>uO^bhscwaV1VK#G-cQX#baon9qSo61rH7Euj3p;%(Y6xPw1EAQW~=hSri?'
    '3MJ6q+Gl6Xlz#=yqFfRnNWl}53p<8il^G$LPdTqTa{T&}l}3{=9H_XHBBMl;D(lrQIrY2Ky7O!}6itmP?|l@8FmYsfnS7GK5QRxGn;d#z+4mioY@Ahw'
    'HH0;@F{!~a5*igo6-@$VqM}&y)t$$k_SS3%#|&nQ;zWNfA>BkIn8e3!_T$j+_*#DakAHWUW=epf-^--{<H2&IxOQLK<_XpOqN*}FlIScgQ^#Lj(1z~X'
    '0RNDzSHWp;c;@&u3ilF{!Z}y9M2&7aFvtj6CnQBQu=d5-=K&{AK=3dtmNf@RMGq=;D@B(XDNdyL(msTwYfk-rK7GUd?v_ltc%+JUk#aF4xkw_iJ5_$2'
    '^u*qDlx5o$U0k^av%FYl#xiNWOPvgQp~MUj(Nz!~wNK8r?JtmL!&^wB%M{azKr2;bBeX`5|IqyI25o)?Yo=?Dx!~;$4LRGJ6d1UKP{gmz)$&f%?;$6|'
    'm~jJ?A_D4e1ot!&ek+o9Vt`jfX?E*<iPNYiX*{Aw0oRJ9R5g9zqgE1dG};R%`%F!r`IANw2S^qFmeDdq%lx!jL~JrzG-x`>Ga2(zW{Mt41xF^vln4MR'
    'QvPpWq$gZ7K%~_2W2sgqNxyw@V<&xp^pxH<`?T#(KK9>xV)W?IF&(YM_jp|S+vpwuTaEwt&2q6%nF<J?oS=LGTFydaa?DUn?SL?pL!zaJN`~r>td_i8'
    'TBg^BY84?yCQ30*$OusrVq}{XTmh5n(O@+cU3D^=vJ(krV<QYB44Y!GvC)@ABj->Yf!y6=`FM*QH>cwlO~Tswi*>U<f#r)`p}SLxH1u(1R#^v&GdtB('
    '%;({Y_5$gzS!<Z*b?3|wGN_4Ph{Y9Xu4Mv1MObG<c<To%!);TBB}tk6o~TUV^(uZR7RNg6G#6)ri)!dl96#3iytpZZ#cir*)tTN9_HY~8zS}q1d}tX*'
    '2#YHtG``%Fq4Uc4ARklBj!Voj(rekfm~tp!F<vi;DfOn=yumBE^AyX#hQK3M8}7v3a|Lu(e(SHrx8_>xWMX~+PFaThv)s$&{@fo=lPsH;3{w-PCQO#@'
    'I0M8$tiI?0liZCND9lS;)$S22Irp(NTvH0Ko<>GO5{YVo1b*c#S`k_=p#3+$&VE2ho5wKTVrC(3WtSFg7+_^wn|&O&gh13!hGN16;f@=rKcSu>h?#@r'
    'tI*b}2vE%dqGO2%L{ZfjCxFQ{s^i0AC{8d*K=}aS1mSZKh{})w1Odbx03$AC5J?a@2N`Rwp(4$R)brzY$pJJYa&|`0rm?llDIaH=7&$vJX=rN1NFE0b'
    'i{o?9WhTu4iU4W`a5x=D08N0ocOZ1*a4wM<%n)quE$Dn3c8xuwE|9P-16BmAqJkGQta|imWJ+wSinvCxIQgvL*N}l2f!N&pQ)|EKkRfQe#oTb)6TSVG'
    'u*!zU+ZQ*WMLtrlpD*ogcnKlj?X7pVa!P`)g2_-Qu1R8}QLzX`V08zauVM?(gw=PkU-NbiVa*-v_#iM<WI}v#4SgDYa81zq4%F?!6dc(XUtAsP@i}2J'
    'jNt4J+y$5!9RL=_i2}bvfjoE-Tt{x$7nc;WRUv0Is-_l{3~`(WZaHo_FgmD1OBlhL2Mf<owCj4`p)?ah1H6-rt4bACc_@Oq9um0PodOOBAO7|CKmOP6'
    '|N86y^TVaL{y6a}(msP@Qp}VN;7{ol4?-r22c71UWZqQd43@b;18@qag6OZ9AbB0t99SG@WOo);SFpeKF^}A=!#Mm`=a-YwGDOQfLd$8}Hr=M1C#YgD'
    'D_`_Hwa-3|Ny>{4LDLmfvmf1XKYit@774W~JP(B;$pzOQyi_3_S`)3kQ?WpMMl30|c437R4<+fG57KpvArOkN9zfsv1k2A)e!7GpO246i{fikQDKq1<'
    'i%Q^BU!_dwn3zD%7x&<LklYC_XFv)F@>VGNk#=ug2QYh0Ygbk5CKe}V6_!i9C47#q<5Ni}P9UAVw0p1?+4fe3IHlSFi&1ebv1qSCh(|G`M@oKkKuty<'
    'wq|J29$GM(DL;QG(8_{wu(CL0Ito)pHC-q^fkqU=eL%o=%U>t8Igk0Lv3S=%Nf$yS-btpEWh8;c+>qa8bJ&9|^J#yB5xAbta3EhU`DJFvZ$1>~w3(pP'
    'Niar67XU1xaOnR1ig{|c?rrRR+wU3(r0jGeRzh&{a%#?ODd;j1)8#ThnWXy&K<dR$`Yc83V<N}6h|&!fvquTIBQU?`bK@eUIGgS8+EXi;Q8Gkn@p;3)'
    '&FDIL-vB^G*lHQ;OxA{{exGc0R*#59c_0)=>kOs;f`4XT_@Dkzd=TtiR38%pDtHu!;!hdx+6y_e{IFu9h#kV>!ov3*OP+~MA=GD;u0l&e5i#{M#HvD6'
    '6JjLW3cdb}Jtb(jC%$1P6ipS!4vQ=5@<jX8Z5y9V``n;C`VGt(n0aNE{^y-cSP`JisS!RpcOny>0O~5PO)QT6%79DB>ep1ie(-tU*-n38;+16%U4~PX'
    '4PSit)Y|ulCL>f4p@o;O3b~4T4?DosM510NF%nEIBIrCrao!T#`-bMihv6p3S)mB`za0t~#7~94VBd-m!tc5HrRl-h7}9+Ei<gv?heKI^O73_|vnKL2'
    'E{fr__t}Yu)G@&D=m;5#-oC}NJkbp02w^S=Cs<%<Lop~dF-dudpr0Uk(1C~IPn{*-`s?*Ap{7)%hD9$&2d|P#gU<eTO8gtcEkBFwoQEsh7av}WxnyO8'
    '37m2${?wN}S(?w0S|y^O6^g`+;x&s2;73wvC0Q90$zpN#B2S+sZN*^k7K-zBQ536(TpJQ9JUpL-Vrrt8vtE@H;kpmSnM}K*5w4^R?**ZV*5EJQ&IKL&'
    'VC-QIPa(hk{bPy>a(XC^74NPszf?1>zSbuA4uxXMA)fm#-Y%qyLfclZi)4uHjT{UuL$U3_abV~-n=D9XlXQfY#czNATS_G8C_^zL(hoh%wR;Qp;!~i-'
    '%HRG@bc_tXlA(wzDn4Pl0RhGaVszNmE@k~m6*n$E@x*ANHgz!m2*pizv<C;;0Dv0-#wP&-%VX2_8B!5ajl_vdzXSfQyCkA-T}9jAPz;H!LFa|X8a)dK'
    'NiSi~A(O5@<zf-E4o}T@Fxj>jGOcD;K~P<h)V=n*q01^tswk6+ZMUIkC-C3)8XxDy?A8+Gy1}R4BveisFm)`rCe&mV!sYnl!RNGbfyA!eYZRDX`fvSl'
    'mwcO$vCF<_=X%qm!j_8$<t8In0bwpwL1|Ix`FE18i&aIGRj{%ooLt7axJemK5l*e(_!y@Nr&n-tj5CBYD>yyEp#d%xCpiD8G8q__<Q-%yz{$uk8X2b6'
    'fa4Z9O(9Q$=v^V%*lEOcrZ`2oG({&<K9{vuzYV8DbIh`31w1W*G`;2mKG_Nx5&G&hrSSrh^yb0H!4(A&uI6J8$AHH{J%kA*CFPcMyiVt=eNV%|Jxhko'
    '8nYo&Zb?r74?QX_5Fq3584{XE=w%X^A@h<sfFj5c-rU1ePcjk}k*MyH@S~9tsEI)R2!Rv$ex?-aGE#8G1q%Kav_kuSKc>Kl0=wIXKp`6kWUr~X-ih&N'
    '#Dr!^(Zw%2f<e$}r2EA>oT>+L0hmirrp#3rR9*P$AaTOpn5(X-y7twb3xH)Z79fluOeayLw#}g_;2MMo=vSUW3_;RBR3R{1(0b-5P6NeRf@)812}EQN'
    'MUWU9F2=Z;`4~<g!{hg(nsT-}eFzUB!k#QQxP6<gT~X~SOkFd4n<rgU<vN*(#)>p8OvM%4-iA4-pg9*JM#RJM8Ed<+mLvcauea+B<<+o$eE{6K#%RPu'
    '7%>-35rVNdFD|0}42lNM6*v2v%cSh)KdVS%CTPs2)o8TPgXy&QqU)q8vd>fu4>qWobv!*NPBI9sJTwbF9kJ?LUX8|G3Ys!VcWc?I4}KVeRKir`G9W@H'
    'rYZ|kU341XoGc?IA*4(J426`6gD&T5Mx>5JQ2ZSeaoM?ZDhyUqh#ACUbn&F<Ra8E#jwoMcStLx7i}zSkqJac+>Fsf31zPXGIx*MDC<vk;rWBIk)3azz'
    'XuZ^av)DKzI3xJxban40ODNo#RJ=CV>&flM!`X>T(fI$Z{nSPk9j0Sol$il-g&JqP6<w`nS$$j|l9oz%Hv}gfv;wEuCf?)C%K+~E4CdXCc@)yAh$N}*'
    'z2I<ip`Lxd=>BlF{YrQNOT2xy%oP|HS*g&4i!G(*>|@#hKCMu3vNGqwIYJ~vTC1h5vpys9NyU5^u<!oW!FBR<eCogy4<j!C#Y#jEY$T(IqWTGn-hR}R'
    'lo?4)B*USo=BBC{?8#$g_RCCia#Myes93amN18Nms#+sYkyY%T0N==@K_INFXbX?jl9OciIKp}H-kcy9^wAW?p)`HCo_u%`aqh#R87>fa@?+>nHYJ*l'
    'MQwg!N<zjaN7LkkrfJZ0^)oaR{S%M5?UN!9GSM3&j&JJW2@f-ub02zES0W=wyTq`y={$GmvnNrRoCGHld~V``>GGlRX=bXpx|FGNTC~GlaeAG=c`{Fk'
    'XM|`qnR=~L&Jc8SdP}ywGNMXpRW#^Tx!5U3Xx9(>oKl@pqS3Bq&m<SXZDv2lW8EF4rp#7sO%vdXarq>KA!uIWEXvewDBkQ;elmD`Y9)AfeAaa2*YYQb'
    'ddc<En0kIV*R)sb_{udK7u-wKPT}!Z@#>R3@Q|osrO1iWE~~2th`PXXBbsGhr!;-`XDZ&OX_vyJpFF3@i6(!5ra$Z_RAo+86V*qC)IEhyXqGw6ifBGE'
    'qF&oG8B>!{Wi+PfCk~RS<CJ_jwcQD_1-vwc3-$;j>V_4m)m1)2Ob|79DZ9&jI0fF<y%-h6+b^axRq=sA_RoAim?8=bm++Uk{P$KEb+3+)JmJgn2F@yB'
    'bGi;z+GNX;@wn-+xOl{(dxQ(lgr913-ce4U%Ec-r_Xco;Kz{bf5Op01=JU&Get9)}9+l2^qWHjNd-xbGUzjPE-+dHF!*6!YtE?44^DWBdZo{BLZ?*m;'
    '10|KuwV<67-z#Saz>zwDMY*MCK5<X4s#H#|CVKTs&r)wRpNM5N>sp(9krRK5#kun{Fw5+-JN6;AJe{Dpn5=M+ktME{&2zQ!+Q(ZcCB~V_Cx-Z0EF!VZ'
    'ZP<VnFssOJ=7WTr_G9rTEBVQuyZ&5Pn&ktkBvF)e-J6~nxckEqk7T9H*D3L|SiJYu$Z&^`&L@T#SS(IC13xQB$|ahDhe(tNRK`9rO-oys_ldbP*LABj'
    'IzPTOEp5&3Q}QRRqa_&67m21>(DFVlf4|^x)jS_A;;ONDhgLF+wp$}oH1hg^kx!Sd9ryewBt{sE_f`a5Al%+ZIVl?XdH!{dlv@^_9?Q>rL!)oxL-f9$'
    'V%PEwYH0AyBkMk!NyG9@WN7$J^#HNJA&0KzlSxc5s~(wvbismASx$_mpn7BqXdgj(o17Gle)Y)crwi6z5IHRx{z@;4q%*u#2wQ-3jmr0DMKe)7uo3)e'
    '=5|3orFE25u{b3JMtOxhynTRlyjEDw?X#LL!kWaTAh@Sg&J9=_;GI|JJBbssjKzc9Bps>5+#^QgO$H%k$E*9iPsQO`hmm3<ikv9&NYVR&M#3S!f>#`U'
    'viGPqp66k9=j0>>krekyI!oN4ceZ6GByFY6K)T9FOCl{7NjnX7cLDZv@B6AF!7?YPh@e{Tt1TS2@ZbpkkkHA^0rY(S=(Jo<PE-?7y-3tq_jtE?%mime'
    'x?gHb&g<JaS2<ZjWX&SkX4AFZ7Bu?yZ_WbgQ?XO!CkxuL^2_F6*`;{o%e;Y#8mLwV!$&T5E~4j~-R}1Pd@4ZM|4D?tvz#=eAz^yDaD5;}7g%B+ww2uk'
    'GBShRvA7*i9}8B$G6stpx0w+6d@>p%wme3R!|m5McFP`5r}JCbI6^Xolgn!5k3XFsf8Hg8iuYQ0vC>foN^(owSmGB}m!7#3BA45|!#@?5k5)gmd~))a'
    ')o4;VZE(&Xr<t7EmL$tb(k|rkagqtoSXQHr&1rpb{W!_wW|vCj6ltD#^)nPNv5&bd!%51*;)4`3F00D9EKc$k79XUErmmi!RQc=(njP^Ivy*Z*c6#z#'
    'LeTVxpO~JU!BK29o-{+ETE6^S2%gy-Ln<D>@l~BKAV~u$`Re@&0kFgMnlo^mUSY5}>Cu>;`^HU>%auJfa(XmuQm?*{T)d{$URvAEMLG5)`DnqA)XU}k'
    'pgB6ugw3DDb<WdpMeA6tid4%7LY<#6iUwcthbbnqT2+s_NE4yt!vyIxu=^}uz$D*tr5`4komf?VR+E&=*QMNyA_=>JYQBs~+~xZ5xyko#ou3CJ_j3LC'
    '<YZfy37xZen%E{+e2ivd%Z~To_^zUtAuK*R?cvCV#hw$aiD3OO!TiLo^9@MvIas`VM&KC+4jQMhAxEbxy~*(TH<?+IH_LoM^5SqL+I^mnrlCnP>QS;^'
    'Ez|r&6!U2}G(W~XO7Q>|27z9FKGGW!pMOJ|qnCYgzSQaUhgXl!jdQB(l4Ycr7}2E(s>ILI&32L{ou70SO;YvnBtiV?c;)2BQbki#nLAg`>})y7!OO5b'
    'U#AsKS5@DkjXMdiGq}x0D*Q6n!}xepRJ>0bqQ)JZiT=074t3B!kA}CVnM2nPG<mR8n%ni%Y@fSJ{}Qu$pDY*SAKk`hRb_?9oH9+ED?Us)xm2o=K7^CH'
    '0gI0i4M~5zmYiivSCHU^#rVLK8pGN)4(I;33n!n8vFVp>{?t^}^@x5&^sAqtKl6@Pjb5Gv$wPvV-Zk>4^ThU`eQqw@M|hQ}KFRPXA#$>u$nu|hImqUY'
    'nVhO5s!Q*=kxk~Frk%Y=Zq!MdI1E>E|9LnNOnRw}RykRkFJAo|+4NepbhE2LkRT)C!%-a`di-F18NKv*f^V4sYbrIOOal>Z5bseQX^<_=eANm=G|d7{'
    'E32d3$R{!vaT~*V<m9?@HtI4Ba<prFgnYXFTq`Nltw!sOm~NGMuf*6gL%+iqfj1=Cg&*Qs{0gqycyuy&_{r(hdRD2^%S(3zXEBJh)i_<vzS;P0=k{$u'
    'MN^OZP4nrB-RreGUR+MB({S}MB7DMd=`PZFx-93SeEd=>E}1vcS@C=W2PDr+FqtnAVsV9Dy_hZ54x}-TbBxcJ#bBp4Bro+>HJ5n{v<S8NCt9GS-6$!F'
    'e5Vudi$(8$1wok6WYc*L_@J0BEeiNZ_`^u$w20}&;v%i%vF|3d@^wY8Bv`Z;!}q4>UeC?(azgJWzf&9@GH<7NNh(_`&a6ww#&w1mnP8vWZHFd5VPvjV'
    'm!w$5;zJaZt!ZSgN04NHfyH<!?)nWTc$vOvB!{*cF<#!8kBW~O^Tb;>5Y5MtNg^sA?BwWNtxT>6ki>1soZGlZGjy@a$<f5H#H|x-hL7Gsv)eBnJT;ej'
    '=Qu4KKD(n)1+(obI8CBw1Ao+<9CeP*Z(sKu6g!1%tKSn+<+LiIl{&eHjMYT+n1pqXq~Bd~7`*A~5#^dF*SNMjP?HSJo1L8Cjd?7H)0|L$E&Pi_&2nvp'
    'k2^-iINj4(p77H;e~FP~sEb7(I4Jq_UB7cq+e{eXNsJuJ>Q$9LZb?N`q|D<qQ!cN{A2K8v(PHsIlF4adNoE(KB%Kcy!!@`W?9u#7v)?B<+*xtB!i;eH'
    '-~4*1RW7Vcx?Zpd&P$2)cw|Lm*JA(;x3OGUd2cR>4OsRoYk<@ia$aqLLGGEJTDhzre=67FNO}p!VsPphx+nYJYyT3NoIF`5c3YP$r$*;qC1+>Q&fF`r'
    'LW%uy8hN`x4jqIVIJ%evQN3#E=fgs8+E{#$KBgo&fT<VoRh#B}6C23a3q5XQ@zD=x>^TeZXK!DdZKOTAxoR5fx}0aCKuLh{UOWvm8UV7E#o|1TV|P2a'
    'xb|R5Hgp3~%MUOzlf~ka1zGx`IDB77(_ia4>!YHFbS&PddWL9Iky>^RDi#j~EY5%Nv;9uXXq~Rt&L4uR9CeE<8L=1-)aT|5&y=w@n?4KknVv89AYeIB'
    'K}3ZQ40i>ew@W)%vbHnv*^=v5#_5`3`OASKnG}oeNTHYlO$!N{fL|S*qP5Cp07Z8v7FP-RQz=jyjaWsLs_`i|M<|X`JsKsNLtft@lz&>rRw+TF?TR7n'
    '!Gc+yU30x#5DQGq=%G~6>mC-Tyy~Ta+at+bjlA!d><UsNsx*os39x1Rgm>8)Lq&$*AU^C~N<q?ROYtD*=N^DOA1N|oiioKPi3NR?XCzba5JfhfO0kF8'
    'Qp+WIVp-?NrW1>#@me#&Or2Mh(sf3Wc7rQEMln$qSjv!kA$mc%e8Fi8Wafs&2Mof;44bL*H#$WoWmpWHWz;`RRy1DHw{AGPSRR(Iyalc`f4WelaK~a;'
    '<&;GrQ`JovR>nPLNZ*@M)@0*?#qs`WXifM*n%|geq%P4U^I`F7&scDMyLR4ja05q4v}Xy1{K?79?5eGDnw)6z`!roo1LFhKR8EaPojvkm3F#(*S<6Co'
    'BfKU`k36e+LSYYthqX+e(xmLgVix(vp1X%&3-1lt#&AY(HVdcsX3dPu!8FjQ5uSk4OGub%1Q!GsbMWhBhRyc&MftYUq(;W#v?~dv)C^na7n&w<7mIf~'
    'H8#4{`Msv;?#JRBxvRZ3n#*b5Y;%o^(76<)$&?0*clu^elrE^XgXvT*jQ9|NR1Et%kF^6?5-r<lj|hMCFs5^{UX#rR7ULc0Pe#YDg0r6$?}t9BF3Mc`'
    '($WwaGf&3v5!Ve-=c217I}a>c7HtFUY=n1*)7L&&HRkWCTE$Rrr5oPBgY;mkaynxU+i=jovwwUp4hP2dr_Ae3li`3iQ<EC6ny@@J$ccppTbK_JbH(V7'
    'Nz7D$8)DM1xb&>W*sF`2Q(T+F$Sp6uju@QX>2Sk2?_2Oi>}!0)K}U;RtK-zco`wKVdo=2&;^wpm$0+b_I>wBVRw90PL3)Z;`=8sf-!|JWxN?gM#@ww$'
    'Dn`A0<J_dum{C|sQ+++D9|D192tx=nL0CQrI@@K4iV)QVfp7j<{hI36llr4rTDF=nD)dQvw;7Bs=&t-8ZY#$=zcYUd+l2x=XtGc&RBF^SQ*j}1xQjoH'
    'li3UzUJ`yK$b7hnW=K>RL4t~K>I5_Iz+?J#M0s_vIwK?-A=ALw0bQhU?LA$n=*4>A(L(U-GzMohE8w5Y(st#uEPwm^Uq4QxXamJcgbDRUdqn&AL5Gv2'
    'S9Gsd_BwPfr;W?j4mqKSK<!nrRP;xR+y1CAf?OtzPv6oQQJZ8%)wPcJnzT_Nm|O_53p%f+KMtpB@xgMmdssuK0!z)uNDg>%|2?7T^3q{mby#YY^4g9M'
    'd$<2EMgd9zR@FLCs3_OelM{$3IP@!AMq!iNKgV=G7RO{*^Pw0x;Bf4?DAq7McG;gBXD<}dT!s+|MLcdikCQ~E2g~B(HGzhNQB=IMRJ;~@x={XWwoF(p'
    'jXx(s(MP{{w$ccw@0{MtC;b}*{sfE>D!{Rb!5i&xbRd{2^RCQ~-(uK_U<Kf0Ef;<N-+4w&6E*!4)NlwLJsLG$)}eQym@UD$0@1Ss0^A(0gl69Dg^FnR'
    'og1-J1}!gEy#+tk80|yU;4qIpiN#4=DO^8}C3{yHQWH`yBV#VJ9Bc?Y6B#V(U#9Ic9-)j+wlA=K7!KHLqj15+u=di$U~xP$70)NzzweeYbLrf7H#>=t'
    'RLC8R|44zzn^SQnrG49XyQZ+r*v2l9Z}u{~x?@n3?*h`;$7>#m7@`No?s$QBY?t}dhx+uHX1_S%bjnK+qvB98<OaL^(m`zB=kBR!_J{5WZ-lS53q{*z'
    'a7J*p0Dpm5IL^a4!TCIVW1$ORL0~ZtZkzMXe*N(@i<g9#3;395Y7VamujcXHt+h=5`8waeCcK`<!~NxgrtcwnJ3QCU(J@5F%+oo*AvcSy2wTl#zc$}3'
    'Yds6D30%*E_sxs*|C&KFT6t_?<++QXok!EcV+#w<Wxw3s_U-JbF<N?TVd>fXVcXBQPD_t1EImBD=diTw*ut{Az09{t3y!I!1^+bvwY1ilzPHvbF9edw'
    '?4o6~*x15i^KWe~tZQNK?}d+{mB$uV-qx}V-{<!(qov2{rEUzCLvH5roE9HnSbQ%HgumW<Ey5fPT7sM{O$T0=vv68~oZsof9GsRQUs!_s9xO`du(SmE'
    '!V*k#Le7$)wa6FNV)Omka%*!~T8w;YF?K`uef-@nejM6cQ8B+iJ--Ru6XrqT{N4Uz+1`t%!g&)fOm}SnhYhzS{1W2{`?@-*E3>+KT6aTrd0uxju3IH^'
    'RaUo}(5(`>I<GsXE7ET()4F2vZGtCsSzdQccXN7`kzvv}5bmWd3k~nW|GoRzVw*FwBG)?}@=gii43;;iU-l%AKrB#XgQ0Tq+n<x<uA5uhre0_h*tXl|'
    '7MpELh!KtIBEf2>F1lK7-^CcmG(0H6V)d;oV_rLnaTKoIvD?Osj_@dd8k`cC1l+A_#YYks97&8@@d~Bl`Rdk#0|IZB4M~|ZY%>D`s_}jIO5^ejW+B)#'
    '#3FYq7;=vn&!DZeoIWp$<|)?YEY_{6NnIw3FIF}<y?_D%df<{6dsg|mNoPPU%)oO9gBPY?Z^tC$EVH7Z2*Rk9dv{=62`NIK>?b+ZMR7hS1H0dDV#q<~'
    '%jL3I?YphZiskn<J6yf=aW*_ogW;%{y!K5Rl^^7uwuASd0~ayd#L%Z-CKkR7BT<IWx<P6H`5MwU@DlYY_;;63%EeQ&0IvS<*W$y*9!aM{zN_+MAxod&'
    '@WyI9zc9vaon}d}lHSSXpm3UE%BpEGTFD^YYgHvSHHlK$OT|SHi%S=azxYUEz9XcFkXa+-lHjeAJHpv9d}=Kjg?^N6<CKd-2lA8?e?X<$9Wl7|rTIch'
    'sgcCBu{8V&t-(Wsh=5ue4iuMA-1SkG1!PNLg+~#EpdYs5#HJK{=pl?><#JP50qgj5>vop4mZhiZ$N<R~#ssGq{mL<EN0(}E*-#PV&*RW}6}Y8A3c(j8'
    'jR-m;Bbd)g1WJ*jvBLw2@SMQyEC7$q<@v>4{1DT!`TmxslOk#rQRDv%wInypOH;DVHfVoTlC}E;x&#-a-6zR$iM_shl-)Zf9xp7W6ZfvImh*CBPFcL|'
    'ndeh~hhb?zv*0#_33p+zibVU=3%%+hVpB$`pGCul<A40$(w`5Z7PEl(LBWK7<Ajm7kgDba@61lTy1jNYZ_@Tt6%PbeIBM*fmRb#OZcje-fUn7@gJse|'
    'mUoa-Etb^cc`ZS;cv4H`wItONNiCVzA}tM*NiCJvB8?1FNiCh%BGn7iNiCDts;HJpYE^lynrc-^t$I|e<ndlNjt3c~h!-Nu4*#-wN)CB7(U5s3)AvI9'
    '>ia%KN5?oHS?zclL%mBwQ5lmJ>+JmfCFORX6+jx7jNLEx&9Ql9)TN2yDd?NLcGT5RK0z<lr*DAR>dzhRrzF9V+seKI<Y>Z&!}d!se0kXn?K6!)J#Nft'
    '(SUW@cd`DS>R0piV=7gso_;)^)OWXS455G6^hbPRA$T$t>`{0e8q0RX-qOCzdJBSHsC`O0tII_dp~Ic%8&gJE)pucO|E{krtZzzF5TU>tI#Uh^F5?qu'
    '!M7galJg|4QixhnVkEV+N#NMq-&(M(cufxDjlC1%81czRXE#P$((st<7*4Z}zqB-pMrVg%UX4>Vm7RupH9^&6b{yu_Xbwz%9_G~)RpZ%#II0$+dcx6<'
    'Pa}R$!K65+5q#G?%>WGnt3_Z#iPjt{$py}4(m%DJ(}Z;%A29q1Xf=Hi%MS0nW2F5yWp;Y!)he3eIy=7eYBg2M&+n|-mKIE1WeW!FM|>?>FBQ+$OI9l`'
    '7mP?m%O$HNsFuoBOIAx#EtxHrtd^o$B3mnYty0Cb>~+e`pit;*X`;omS;}Y`s>QNN%4$_qtIXyot3^{(*Vz=MwVv#y5yS&#(&Tt9e_xknAA7R0qwzQe'
    '5j6Z~tICw4>{u>-m5a3H>%$#McT|>}HwfNMkzP$tn={qdBl4TGuhh2m9yFWY1nuAQw<HT22v#q`(nqA1r%v#Y$WhppI?e}{hSR<T$cCCJS6-`7ixL}3'
    'zNuP4O}?tz#*<u#6`(iXB3v>WZ0%7Eh2sEwmUUSNg^1AOIeI=_L&fBSW>OM`yfxDgdQLhf7J%bv+M6;YB@{lqU)rnv=_M@5LonFG{+mg@=m9b+hK+KN'
    'mF7funqAtP4utA^TD@hy+PlmY=miWRl4XAuJS2Mrp+nBx1unv_8v&x@(Zi0K6{Agt{Ekymh1xwxToJ+2!Gh;<637y>{jl{eOh$c}9S||z58jaCVp*2%'
    'EnsS%4tX`nsdup4YIsICH-1cZ(h8cKE>%4--iSk!$YeVOL7WLy&t1jfU*W5F&omHfGGo!Y_^r6!Qh%m^eragrt0~-J-J|R!4q<O4@Ll_S^XIR`NY}G+'
    '>ZTvA&)dcaj}IG|MLJKdC*orFFA!NUz;?zRTt`}mfrRKqS_jA~+4h@0rBMbNe6ypkD(lC-ip)3Z!+6>meZMm)je_$)M#S-9n34X$@rhZ2fttVidb`!}'
    'HaFq*wb?&4U!8shgum_rZ)$^cI*f)l-U<*&99(?C$nXg@v&kVg=u^<fzZ^*eQsx>E7gXC<A}vIjYa!A~q=YDQB}7_@bP#2(gGeip3Zl$e?z9qVAj(_='
    'kyR3bhpMaE^S{IG>Mhr;&<koV&~MqpC@_~t(pr?c)*_`$N{cdATBMaoXHn)li?kA{EXthaO)G(PPREFtoC9g1%9pd(<_V+X&|rp-CM2{9mitkscqW#*'
    '%2{e>TkI<lM_rq2A*7W^{;G3^I;}*~SDo|IX(f`q>YSZUE0N??=iGE!iG-OtXOq)PB+Aq|mz-82L8i``<g^m;$W_iGr<EGY;}_YgB$Y0zADq_H=X5Ls'
    'kka=F`3R5&ASLZ1Gp{uu4?sac{RmJ5AZ3jd<`JL_K*|;=sz-n-04YbLs2>690HoxQvU~(+0+3QdO7;k_3P4H)DfuJ7IsgxG=`cS1NFV(on$sTt<KN@b'
    'ImESt`4H|R<JviJLb6AYaq%1kA^9W7xOxtfkm3<!Ts{XyNcjjduAhS@q<RDy7tp~FQa^%>E9g)W(maBUOXyG&vU&s=)=={Hp=OplNRICK4?L3vyfVT_'
    'xrzzvEqTCDY!ShdR#@7~+$oh{Vk^XZD87i`Nh>^SMXq9EE5y7gv4{{!D<W%U9TQt2?nf>6OD3(ztQEeBiLDU(q?X56C9SBe6}F0rtq?1v^ddqht>~<k'
    'aupL>Azn<GMTALOG1*GytC-jdF>|U~L{v#DRo04F#l%*K8&vfoqE1?=vsTzDCbZJh^-)e|f8@%M+#W4m9;|!-7dqL})xp>UxNvi{ba62L0504cEnOQ-'
    'Jb(+gMoX6llMmp+ozc>j!5$gUaAUM|VKDswF5DL_T^GzefD5-pOP2+!9>9gWVoRKAIh}VesE?S`EitJjyMqbqWJ~;M$?sso8rc$ST8cZEus*iLm6q}j'
    'CajGuF{CqlCdY(zu_a!#)ORppO>Bt`EzKQFSPxs`K+Eb5Cai@mF`s392NTvoZ}|<9LK9Ub{NnYMBh4ZFav=?t$uZtQ$yeRevBaibN*%<KGigs}5-;z8'
    '!b02AfyCK8P*`buI*mBL2MSAVPe&0K_dsE-?dcri@*XHGwmls}T-^hO)wZV-i0gZxu-x`^{BUy*6xQ3G&K_Rf1BC^5rh{kZ!iie6(b03Jqi5#28Kc6Q'
    'IMd-XbJ>hhVNsmv_?fwC#;C9=&U66HTr^`;SQckGf@ZFnF)FN!GaW)Rm&_Oy7RH&5p_waYj0!8`Ob5}-1v5s4rE#XCXlCD!QDJQi#P;av+Q2uZT!I;L'
    '3FZzctfqlP7|oU;i3~GjGK^H~px6!xF`6$!5-DZ~QVfepGFCvb9TH=-ScW8W%#h15^*SiFLxPN!%aBBp8FEQRuY+PcB+BTOp-yC(A(v&Ubx>@Fgc-fk'
    'wM3d3a%sk_gJL@*&gd0LBF_xDJfqh^u^kd<^vVz<63viHG-e$X+aZxguPmKJrWtaX<_;*fLsHG0w3XPLjqi(r<eE9X#tbs{lq8$keLh9<h@^{SZpNla'
    '5s`Fx%+0|RDI=0D5LrEfR1ry+$lQEPkvbyjBAL-M%^*!g(q%HEJDNdO5lI)yj6P@v8Sa!9;y`DPIW)<hvY;=-g3eqjV^ml>FT{h+i0xxkST`@kgw7}~'
    '#;CAnUWf~w*&1V1ST8TchR$q@F)FN;7ve)_w!|0}*2xPoqBGlJj0$Vyg*eff1wBTE_3=WisJuVRer+V|1~0P*GH}uNb%<5qEPEeI>d-aMdG<V()S)Yy'
    'i|lnQsYBN*&mYH<I&?L1mA#E6b?92-I(r&R>d=+MP4+UD)S>HySJ}f@QirYrp1+GFbvCJK@0zvj*ATxi-W<cRxst5OMs_OUwKp>d8&34_Y0PL|gOSc~'
    'N=ymOXkw`nJ$zPrG_zER9zH8Inp&zv51-X|ZmAMId{$yKxm1ZBJ}WkwU8+P6pH(@UUaCY7pQawoFIA$4Pg9L1I94K|j_HM)Z*s;CB-14Zwo9O}97v=~'
    'Y*?2-Jb>tg)#e^Z1P~pty1oaJ0YvAkuI_<U0MYSU-v#LaqSLj$JLD#S=y0v?4n!3|bha9C59E*5E7_gKcH%yt#uz-@4*lh$hbpgfMPn>o&KYC5jBi-~'
    'MQT;W`UXc3%-ZeB|M4S6KZ<UIFTXtZuPQoOis+YQ|9<)Q7DKGO7YYq2e(5jIEI1ybcD;3BYEttbs8oEoTu<${j}gTQr5{7R8O5&yh3Gf#zNMrQ&pkvj'
    'It?dq{#4~`!iY%25v+KXCVtgUr4ju4wqTE9ug24*y#Wslv9=)^9cHSrFo<1c?_rfV_Pw>N#@8ZmEeFN7W^;03NRFq2iyZ~&dVKw_rlipJRh7mLit*jx'
    'yPPLy;j07k`)L?X-gi%8{?WD>y3~lzj)JQdtzLfp;lu5M3$F1#nAclTG+XevoY+DB`lA+g6|Ivo$AWWxXnW^MQ8fGSCws$y2lCEi{`VjMCW37u!^|?l'
    'N&RWx^&5LLHOC%Iaq!ssm;WJE9oDF^{fcP~z{DRze5N6cf~poSnh#99@-l>T`j_7{vS7}su5v~#cH~dCZTL7B4F#+m56$_4Yd-ecLWMNJ3KQMc5-3K4'
    'L95vi7AvLlhrc>1KYsn6zfZrtp80kcWg>k0kN+XPM?X4IyC{#rLt)vA`o~}X{Fh(<-+%o1x6onjt|&KgF0!BtQf)UXjkpW7cK8TBE0M%+zy9@KB8AGV'
    'Q#drvS@*+cLgUB(A|mx;MB3+^%3uHTw_krmDs{C^1=B>d35kF84{Z=fo0pCo!JKRTF%M;ex5|&d+6zdBg;LAf?|Ro9&h-iXRDtJV1w;&so9KW39g=kg'
    'FouqOY9L@2#v8L|{qA(O-7I(rH{>9OAbJTRQiMtn^%5e)DRR;IKimE299jK3`NMlHS@0*<9DSz202lnpo7R5rsd$B$cwjo=cCvyU(ft4Og`QDWkYxm^'
    ')j;~q&B4QGz|zGT<$x>wVE@oCS3l1ncnRmI;N>3)DuTR0;iq5(EwPGofh85<5_RyPGwdu&U?O)7a)n)Q0XhxW3s{U=!ftpqfga3R0?}#$Kj0O@bO)Dm'
    'ZlQjWo){%QL|<x7Vp2|GUUDMhbau#r1&}+^q1*7G+b7_M6vFl9B^D*-Wg#?!;{#1IY)2s$-KptCNERDl(F>|-&~%oi8AESbge>yL+$aG(;L8KHy0xqz'
    '6_^<_(@No4$pY5Sf)so1;Au|}$l7<v`bcnIf^{gM1{5deAPlf1U_h{=kFUQU$beV?9Pj^Y'
)
_AMBIGUOUS_BARCODES = ('111009', '17427', '17796893021870', '19677', '26328', '29078', '32184', '38836', '39409', '40061', '41528', '41821', '42671', '44272', '45999', '46010', '47257', '47781', '49100', '49102', '49827', '49854', '49856', '52528', '55253', '692452117102', 'BNVX4')

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
        self.conn.execute('CREATE TABLE IF NOT EXISTS barras_ambiguas (codigo_barras TEXT PRIMARY KEY)')
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

    def mark_ambiguous_barcodes(self, barcodes):
        """Never resolve duplicate physical barcodes automatically."""
        with self.conn:
            self.conn.executemany('INSERT OR IGNORE INTO barras_ambiguas(codigo_barras) VALUES (?)',
                                  ((str(b),) for b in barcodes))
        return len(barcodes)

    def write_legacy(self):
        # Maintain backward-compatible JSON and prevent lost selections after older reinstall.
        codes = sorted(self.store_codes())
        temp = self.legacy.with_suffix('.json.tmp')
        temp.write_text(json.dumps(codes, ensure_ascii=False), encoding='utf-8')
        os.replace(temp, self.legacy)

    def resolve_verified_barcode(self, barcode):
        """Scanner: ONLY an explicit barcode→IdArticulo association, never a SKU guess."""
        barcode = str(barcode).strip()
        if self.conn.execute('SELECT 1 FROM barras_ambiguas WHERE codigo_barras=?', (barcode,)).fetchone():
            return None  # This barcode exists for multiple different IdArticulo in the source.
        if not barcode:
            return None
        row = self.conn.execute(
            'SELECT c.codigo, c.descripcion, c.precio, c.en_tienda '
            'FROM codigos_alternativos b JOIN catalogo c ON c.codigo=b.codigo '
            'WHERE b.alternativo=?', (barcode,)).fetchone()
        if row:
            return dict(code=row[0], name=row[1], price=money(row[2] or '0'),
                        store=bool(row[3]), barcode_source='asociacion')
        # Codigos propios Cotyland: el punto es un prefijo REAL que imprime el
        # negocio delante de un IdArticulo de ocho digitos. La lectura literal
        # ".10302830" identifica SOLO el IdArticulo exacto "10302830".
        # Nunca se aplica a lecturas normales ni a codigos parciales; tampoco
        # se guarda una asociacion supuesta en la tabla de barras verificadas.
        if len(barcode) == 9 and barcode.startswith('.') and barcode[1:].isascii() and barcode[1:].isdigit():
            own = self.get_product(barcode[1:])
            if own:
                own['barcode_source'] = 'codigo_propio'
                return own
        return None

    def verified_barcode_for_label(self, code):
        """Return the one unambiguous known physical barcode, or None.

        Never print a SKU in place of a barcode and never arbitrarily choose
        among different physical barcodes associated with the same article.
        """
        rows = self.conn.execute(
            'SELECT b.alternativo FROM codigos_alternativos b '
            'LEFT JOIN barras_ambiguas a ON a.codigo_barras=b.alternativo '
            'WHERE b.codigo=? AND a.codigo_barras IS NULL', (str(code),)).fetchall()
        return rows[0][0] if len(rows) == 1 else None

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
        self.db.mark_ambiguous_barcodes(_AMBIGUOUS_BARCODES)
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
        scanned_barcode_copies = {}
        info = tk.StringVar(value='0 lecturas. Código desconocido = pendiente, sin asignación automática.')
        ttk.Label(panel, textvariable=info, wraplength=960).pack(anchor='w', pady=8)

        def display():
            selected = tree.selection()
            tree.delete(*tree.get_children())
            for barcode, item in session.items():
                product = item.get('product')
                status = ('Código propio' if product and product.get('barcode_source') == 'codigo_propio'
                          else 'Verificado') if product else 'PENDIENTE'
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
            if product:
                # A real, verified manual scanner reading means the item is physically
                # in the store. Persist this choice before preparing labels.
                if product['code'] not in self.in_store:
                    self.save_store([product['code']], True)
            if product and auto_print.get():
                code = product['code']
                scanned_barcode_copies[barcode] = scanned_barcode_copies.get(barcode, 0) + 1
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
            self.labels(only_codes=codes, barcode_counts=dict(scanned_barcode_copies))
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
                if code not in self.in_store:
                    self.save_store([code], True)
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
            if code not in self.in_store:
                self.save_store([code], True)
            self.chosen.add(code)
            scanned_barcode_copies[barcode] = scanned_barcode_copies.get(barcode, 0) + 1
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

    def labels(self, only_codes=None, copies_override=None, barcode_counts=None):
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
        qty_source = self.qty if copies_override is None else copies_override
        # Each physical barcode scanned is printed exactly as received. A SKU may
        # have several legitimate barcodes, and these must not be merged or guessed.
        items_to_print = []
        if barcode_counts is not None:
            selected_by_code = {x['new']['code']: x for x in selected}
            for barcode, count in barcode_counts.items():
                product = self.db.resolve_verified_barcode(barcode)
                if not product or product['code'] not in selected_by_code:
                    messagebox.showerror('Código sin verificar',
                        'No se puede imprimir el código de barras ' + barcode +
                        ' porque su asociación no está verificada.')
                    return
                if count > 0:
                    items_to_print.append((selected_by_code[product['code']], barcode, int(count)))
        else:
            missing_barcodes = []
            for item in selected:
                code = item['new']['code']
                barcode = self.db.verified_barcode_for_label(code)
                if not barcode:
                    missing_barcodes.append(code)
                else:
                    count = int(qty_source.get(code, 1))
                    if count > 0:
                        items_to_print.append((item, barcode, count))
            if missing_barcodes:
                messagebox.showerror('Código de barras no confirmado',
                    'No se imprimirán etiquetas con SKU en lugar del código de barras. '
                    'Estos artículos no tienen un único código de barras verificado: '
                    + ', '.join(missing_barcodes[:12]) +
                    '\n\nEscaneá y verificá el código de barras real antes de imprimir.')
                return
        copies = sum(count for _, _, count in items_to_print)
        if not copies:
            messagebox.showwarning('Sin etiquetas', 'No hay etiquetas preparadas para imprimir.')
            return
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
        for item, barcode, quantity in sorted(items_to_print, key=lambda x: (x[0]['new']['name'].casefold(), x[1])):
            product = item['new']
            for _ in range(quantity):
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

                # Exact barcode TEXT (including initial zeros and letters), never SKU.
                # No barcode image; the numeric/alphanumeric identifier is enough.
                identifier = str(barcode)
                c.setFont('Helvetica-Bold', 10.0)
                c.drawCentredString(x+w/2, y+5.1*mm,
                    clip('Cód. barras: ' + identifier, 'Helvetica-Bold', 10.0, w-4*mm))
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


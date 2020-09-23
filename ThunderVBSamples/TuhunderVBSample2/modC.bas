Attribute VB_Name = "modC"
Option Explicit
Private Sub AsmSettings()
'#asm' Option Scoped
End Sub
'Here we use the text as Ascii 1 byte per char ..
'You can chose to directly modify the VB string's text tho...
Public Sub StrRot_A(ByRef str As Byte, ByVal cnt As Long, ByRef outb As Byte, ByVal numRot As Long)
'#c'void StrRot(char*str,int cnt,char*outp,int nrot)
'#c'{
'#c'int i=0;
'#c'outp[cnt]=0;
'#c'outp[cnt-1]=str[0];
'#c'for (i=1;i<(cnt);i++){
'#c'outp[i-1]=str[i];
'#c'}
'#c'for (i=0;i<(cnt);i++){
'#c'str[i]=outp[i];
'#c'}
'#c'if (nrot>=1) StrRot(str,cnt,outp,nrot-1);
'#c'}
End Sub

'This is the 2 byte Version of the above.. can be used directly with vb strings..
Public Sub StrRot_W(ByVal str As Long, ByVal cnt As Long, ByVal outb As Long, ByVal numRot As Long)
'#c'void StrRot(short*str,int cnt,short*outp,int nrot)
'#c'{
'#c'int i=0;
'#c'outp[cnt-1]=str[0];
'#c'for (i=1;i<(cnt);i++){
'#c'outp[i-1]=str[i];
'#c'}
'#c'for (i=0;i<(cnt);i++){
'#c'str[i]=outp[i];
'#c'}
'#c'if (nrot>=1) StrRot(str,cnt,outp,nrot-1);
'#c'}
End Sub


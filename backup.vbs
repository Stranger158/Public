'������ ��� ��������� ������ � ����������� ������� � ������� ���������

'��� ������������
Const SRC = "C:\Users\n.izvolskaya\Desktop\testfolder"

'���� ������ 
Const dst = "C:\Users\n.izvolskaya\Desktop\123\"

'������� ����� ��������� �����
Const PREFIX = "backup"

'���������� ��������� �����
Const EXT = ".zip"

'���������� �������
Const HISTORY = 2

'��� ������������
Const PROGRAM = """%ProgramFiles%\7-Zip\7z.exe"""
Const OPTIONS = "a -r -tzip -mx1"

'�������� ���������
Const NOTIFY = True

'�� ��������� ������ ��������
On Error Resume Next

'��� �����
Dim file
file = PREFIX & "_" & FormatDateTime(Date, 2) & EXT '2 = vbShortDate

'���������� ���� ����� ��� �������
Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(file) Then
	If NOTIFY Then WScript.Echo "��������� ����������� �������� �� �������" & vbCrLf & SRC & vbCrLf & "����� ��� ��� ������ �����"
	WScript.Quit
End If

'������� �����, ��������� ��� ������
Dim sho, ret
Set sho = WScript.CreateObject("WScript.shell")
ret = sho.Run(PROGRAM & " " & OPTIONS & " " & file & " " & SRC, 7, True)

'�������� ������ ������
Dim dir, fc, f, x, i, names(), dates()
ReDim names(0)
ReDim dates(0)
Set dir = fso.GetFolder(".") '������� �������
Set fc = dir.Files '��������� ������
i = 0
For Each f in fc
	If Left(f.name, Len(PREFIX)) = PREFIX And Right(f.name, Len(EXT)) = EXT Then
		ReDim Preserve names(i + 1)
		ReDim Preserve dates(i + 1)
		names(i) = f.name
		dates(i) = f.DateLastModified
		i = i + 1
	End If
Next

'�������� ��������� ������
Dim j, dmax, imax
For j = 1 To HISTORY
	dmax = DateSerial(2000, 1, 1)
	For i = 0 To UBound(dates)
		If dates(i) > dmax Then
			dmax = dates(i)
			imax = i
		End If
	Next
	dates(imax) = DateSerial(2000, 1, 1)
	names(imax) = ""
Next

'������� ���������� ������
For i = 0 To UBound(names)
	If Len(names(i)) > 0 Then
		fso.DeleteFile(names(i))
	End If
Next

'== ���������

'������� ��������� �� ���� ������
If NOTIFY Then
	Dim msg
	Select Case ret
	Case 0
		msg = "������ ����� " & file
	Case 1
		msg = "������ ����� " & file & vbCrLf & "��������� ����� �� ������� ��� ���� ������ � ������� �� ��������� � �����"
	Case 2
		msg = "��������! ��������� ������ ��� �������� ������"
	Case 7
		msg = "��������! ��������� ������ � ��������� ������"
	Case 8
		msg = "��������! ������������ ������"
	Case 255
		msg = "��������! �������� ������ ���� �������� �������������"
	Case Else
		msg = "������ ��� �������� ������, ��� " & ret
	End Select
	WScript.Echo "��������� ����������� ��������:" & vbCrLf & SRC & vbCrLf & msg
End If
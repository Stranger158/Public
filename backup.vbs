'скрипт для архивации файлов с последующим бекапом в сетевое хранилище

'что архивировать
Const SRC = "C:\Users\n.izvolskaya\Desktop\testfolder"

'куда залить 
Const dst = "C:\Users\n.izvolskaya\Desktop\123\"

'Префикс имени архивного файла
Const PREFIX = "backup"

'Расширение архивного файла
Const EXT = ".zip"

'Количество архивов
Const HISTORY = 2

'Чем архивировать
Const PROGRAM = """%ProgramFiles%\7-Zip\7z.exe"""
Const OPTIONS = "a -r -tzip -mx1"

'Выдавать сообщение
Const NOTIFY = True

'Не завершать скрипт аварийно
On Error Resume Next

'Имя файла
Dim file
file = PREFIX & "_" & FormatDateTime(Date, 2) & EXT '2 = vbShortDate

'Пропустить если архив уже делался
Dim fso
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(file) Then
	If NOTIFY Then WScript.Echo "Резервное копирование каталога не удалось" & vbCrLf & SRC & vbCrLf & "Архив уже был создан ранее"
	WScript.Quit
End If

'Создать архив, запомнить код ошибки
Dim sho, ret
Set sho = WScript.CreateObject("WScript.shell")
ret = sho.Run(PROGRAM & " " & OPTIONS & " " & file & " " & SRC, 7, True)

'Удаление старых файлов
Dim dir, fc, f, x, i, names(), dates()
ReDim names(0)
ReDim dates(0)
Set dir = fso.GetFolder(".") 'рабочий каталог
Set fc = dir.Files 'коллекция файлов
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

'отобрать последние архивы
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

'удалить устаревшие архивы
For i = 0 To UBound(names)
	If Len(names(i)) > 0 Then
		fso.DeleteFile(names(i))
	End If
Next

'== РЕЗУЛЬТАТ

'вывести сообщение по коду ошибки
If NOTIFY Then
	Dim msg
	Select Case ret
	Case 0
		msg = "Создан архив " & file
	Case 1
		msg = "Создан архив " & file & vbCrLf & "Некоторые файлы не найдены или были заняты и поэтому не добавлены в архив"
	Case 2
		msg = "Внимание! Произошла ошибка при создании архива"
	Case 7
		msg = "Внимание! Произошла ошибка в командной строке"
	Case 8
		msg = "Внимание! Недостаточно памяти"
	Case 255
		msg = "Внимание! Создание архива было прервано пользователем"
	Case Else
		msg = "Ошибка при создании архива, код " & ret
	End Select
	WScript.Echo "Резервное копирование каталога:" & vbCrLf & SRC & vbCrLf & msg
End If
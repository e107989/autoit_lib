
;; # Wizard Functions

$TEST_WIZARD = False

;; $wiz_letter : The letter in the title bar, usually A or B
;; $wiz_ip : The address, examples are 'MainframeA' and '172.26.20.244'
;; Return: A string containing the current contents of the specified wizard screen.
Func getWizardScreen($wiz_letter="", $wiz_ip="172.26.20.244")
	activateWizard($wiz_letter, $wiz_ip)
	Send("!{F10}^c")
	Return clipget()   
EndFunc

;; Desc : Send a string of keys to wizard and wait for the screen to change. 
;; $key : The key string to send.
;; $timeout : How long to wait in milliseconds.
;; $wiz_letter : The letter in the title bar, usually A or B
;; $wiz_ip : The address, examples are 'MainframeA' and '172.26.20.244'
;; Return : True if wizard responded. False if timeout reached.
Func sendWaitWizard($key, $timeout=5000, $wiz_letter="", $wiz_ip="")
	Local $screen = getWizardScreen($wiz_letter, $wiz_ip)
	Send($key)
	Return waitForWizard($screen, $timeout, $wiz_letter, $wiz_ip)
EndFunc

;; Desc : Get some text from a wizard screen. $row, $col, and $len may be passed as a single argument in an array.
;; $screen : A string representing the wizard screen.
;; $row : The row to get from.
;; $col : The column of the first character.
;; $len : The length of the string to get, including the first character.
;; Return : The string at the location described.
Func getFromWizardScreen($screen, $row, $col=-1, $len=-1)
	; Function overloading....
	If $len == -1 Then
		$len = $row[2]
		$col = $row[1]
		$row = $row[0]
	EndIf

	Local $lines = StringSplit($screen,@CRLF,1)
	Local $line = StringSplit($lines[$row],"")
	Local $max_length = UBound($line)
	Local $output = ""

	For $i = $col To $col+$len-1
		If $i >= $max_length Then
			ExitLoop
		EndIf
		$output &= $line[$i]
	Next

	Return $output	
EndFunc

;; Desc : Search for a string in a wizard screen.
;; $screen : A string representing the wizard screen.
;; $string : The string to search find.
;; Return : The [row, col] of the string in the screen. [-1, -1] if not found.
Func searchForInWizardScreen($screen, $string)
	Local $lines = StringSplit($screen,@CRLF,1), $loc[2] = [-1, -1]
	For $i = 0 To UBound($lines)-1
		$in_line = StringInStr($lines[$i], $string)
		If $in_line <> 0 Then
			$loc[0] = $i 
			$loc[1] = $in_line 
			ExitLoop
		EndIf
	Next
	Return $loc
EndFunc

Func getColFromWizardScreen($screen, $row, $col, $len, $height)
	Local $lines = StringSplit($screen,@CRLF,1)
	Local $output[$height]
	For $i = 0 To $height-1
		$output[$i] = ""
		Local $line = StringSplit($lines[$row+$i],"")
	For $j = $col To $col+$len-1
			$output[$i] &= $line[$j]
		Next
	Next
	Return $output
EndFunc

Func activateWizard($wiz_letter="", $wiz_ip="")
	If $wiz_letter == "" Then
		$win = WinWaitActivate(getWizardTitle("A","172.26.20.244"),"",1)
		If $win == 0 Then
			$win = WinWaitActivate(getWizardTitle("B", "172.26.20.244"),"",1)
		EndIf
		Return $win
	Else
		Return WinWaitActivate(getWizardTitle($wiz_letter,$wiz_ip))
	EndIf
EndFunc

Func getWizardTitle($wiz_letter, $wiz_ip)
	Return "(" & $wiz_letter & ") TN3270 (" & $wiz_ip & ") - PowerTerm InterConnect/32"
EndFunc

Func checkWizardScreen($check_screen, $wiz_screen)
	If $check_screen == "MSRH" And checkWizardText($check_screen, $wiz_screen, 2, 2) Then
		Return True
	EndIf
EndFunc

Func checkCursor($wiz_screen, $row, $col)
EndFunc
	
Func checkWizardText($text, $wiz_screen, $row, $col)
	Return getFromWizardScreen($wiz_screen, $row, $col, StringLen($text)) == $text
EndFunc

Func vlookupWizard($screen, $lookup_val, $src_col, $out_col, $position=False)
	If UBound($src_col) > UBound($out_col) Then
		If $position Then
			Return -1
		Else	
			Return ""
		EndIf
	Else
		For $i = 0 To UBound($src_col)-1
			If $src_col[$i] == $lookup_val Then
				If $position Then
					Return $i
				Else
					Return $out_col[$i]
				EndIf
			EndIf
		Next
		If $position Then
			Return -1
		Else
			Return ""
		EndIf
	EndIf
EndFunc

Func goToWizardScreen($command)
	Send("{ESC}")
	Send($command)
	Send("{ENTER}")
EndFunc

; Checks to see if the screen has changed because sometimes wizard
; takes a long time to load a screen. Uses a delay method similar
; to that of TCP where the delay grows exponentially.
; ** TIME IN MILLISECONDS **
Func waitForWizard($old_screen, $max_delay=5000, $wiz_letter="", $wiz_ip="")
	$new_screen = getWizardScreen($wiz_letter, $wiz_ip)
	$delay = 50
	While $new_screen == $old_screen And $delay < $max_delay
		$new_screen = getWizardScreen($wiz_letter, $wiz_ip)
		sleep($delay)
		$delay *= 2
	WEnd
	If $delay >= $max_delay Then
		Return False
	Else
		Return True
	EndIf
EndFunc

; --------------------------------
; Excel Functions
; --------------------------------

Func activateExcel($file_name)
   WinWaitActivate("Microsoft Excel - " & $file_name)
EndFunc

Func getFromExcel($file_name, $col, $row)
   activateExcel($file_name)
   gotoCellExcel($col, $row)
   Send("^c")
   Return clipget()
EndFunc

Func getArrayFromColExcel($file, $col, $start, $end)
	activateExcel($file)
	Local $arr[$end-$start+1]
	gotoCellExcel($col, $start)
	For $i = 0 To $end-$start
		Send("^c")
		$arr[$i] = StringReplace(clipget(),@CRLF,"")
		Send("{DOWN}")
	Next
	
	Return $arr
EndFunc

Func getLastRowInColExcel($file, $col, $start)
	activateExcel($file)
	gotoCellExcel($col, $start)
	Send("^{DOWN}")
	; Return ThisRow
EndFunc	

Func gotoCellExcel($col, $row)
	Send("{F5}" & $col & $row & "{ENTER}")
EndFunc

Func putInExcel($val, $file_name, $col, $row)
	activateExcel($file_name)
	gotoCellExcel($row, $col)
	Send($val&"{ENTER}")
EndFunc

Func putArrayInColExcel($arr, $file, $col, $start)
	activateExcel($file)
	$len = UBound($arr)
	gotoCellExcel($start, $col)	
	For $i = 0 To $len-1
		Send($arr[$i])
		Send("{ENTER}")
	Next
	
	Return $arr
EndFunc

Func copyRangeFromExcel($file, $start_col, $start_row, $end_col, $end_row)
	activateExcel($file)
	$width = ltr2num($end_col) - ltr2num($start_col) 
	$height = $end_row - $start_row
	gotoCellExcel($start_row, $start_col)
	For $i = 0 To $width - 1
		Send("+{RIGHT}")
	Next
	For $i = 0 To $height - 1
		Send("+{DOWN}")
	Next
	Send("^c")
EndFunc

Func pasteInExcel($file, $col, $row)
	activateExcel($file)
	gotoCellExcel($col, $row)
	Send("^v")
EndFunc

Func ltr2num($letter)
	$sum = 0
	$mult = 0
	$alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	$chars = StringSplit($letter, "")
	$len = $chars[0]
	For $i = $len To 1 Step -1
		; Alert($i)
		$cval = StringInStr($alpha, $chars[$i])
		; Alert($cval)
		$sum += $cval * ( 26 ^ $mult )
		$mult += 1
	Next
	Return $sum
EndFunc

Func sortExcelCol($col)
	gotoCellExcel(1, $col)
	Send("!hss")
EndFunc

; --------------------------------
; Generic Functions
; --------------------------------

Func WinWaitActivate($title, $text="", $timeout=0)  
	WinWait($title, $text, $timeout)
	If Not WinActive($title, $text) Then WinActivate($title, $text)
	Return WinWaitActive($title, $text, $timeout)
EndFunc

Func Alert($msg)
	MsgBox(64, "Alert", $msg)
EndFunc

Func Continue($msg)
	$res = MsgBox(4, "Continue?", $msg)
	If $res == 6 Then
		Return True
	ElseIf $res == 7 Then
		Return False
	EndIf
	
	Return 0
EndFunc

Func Input($prompt)
	Return InputBox("Enter Input", $prompt, "", "", -1, 130)
EndFunc

Func padLeft($str, $pad, $len)
	$output = $str
	If StringLen($str) > $len Then
		Return $str
	Else
		While StringLen($str) < $len
			$output = $pad & $output
		WEnd
		Return $output
	EndIf
EndFunc

Func stripWSArray($string_array, $flag)
	For $i = 0 To UBound($string_array)-1
		$string_array[$i] = StringStripWS($string_array[$i], $flag)
	Next
	Return $string_array
EndFunc

Func AlertArray($arr)
	$out = ""
	For $i = 0 To UBound($arr)-1
		$out &= $i & ": " & $arr[$i] & @CRLF
	Next
	Alert($out)
EndFunc

Global $LOG_FILE_NAME = 0
Global $LOG_FIRST_ENTRY = 0
Func _OpenLog($filename="log.txt", $mode=2)
	$LOG_FILE_NAME = FileOpen($filename, $mode)
	$LOG_FIRST_ENTRY = True
	$header = "[" & @CRLF
	FileWrite($LOG_FILE_NAME,$header)
EndFunc

Func _CloseLog()
	$footer = @CRLF & "]"
	FileWrite($LOG_FILE_NAME,$footer)
	FileClose($LOG_FILE_NAME)
EndFunc

Func _Log($msg, $type="")
	; Output log is in JSON array format for indexing and searching
	If $type == "" Then
		$type = "MSG"
	EndIf
	If Not $LOG_FIRST_ENTRY Then
		$logmsg = ','
	Else 
		$logmsg = ''
		$LOG_FIRST_ENTRY = False
	EndIf
	$logmsg &= '{' & @CRLF & '"time" : "' & @HOUR & ':' & @MIN & ':' & @SEC & ':' & @MSEC & '", "type" : "' & $type & '", "message" : "' & $msg & '"' & @CRLF & '}'
	FileWrite($LOG_FILE_NAME, $logmsg)
EndFunc
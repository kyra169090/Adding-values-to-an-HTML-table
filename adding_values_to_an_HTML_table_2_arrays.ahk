#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance,Force
SetBatchLines, 40ms         ; when I did not use this line my script often just stopped, which seemed pretty abnormal


Esc::ExitApp ; Exit the application on pressing Escape
/*
************access Path to G33kDude's Chrome.ahk file needs to be changed*********************
*/
#Include ; put the Chrome.ahk path in here

page :=Chrome.GetPageByTitle("example","contains")      ; This will connect to the specificed tab
If !IsObject(page){
MsgBox % "The page wasn't found"
ExitApp
}
/*
************Press Alt+1 after selecting the first column in Excel, then Alt+2 for second column********************
*/

!1::HandleSelection(Array1)
!2::HandleSelection(Array2)

HandleSelection(ByRef arr) {
    arr := [] ; Empty the array
    Clipboard := "" ; Empty the Clipboard
    SendInput, ^c ; Sends Control + C for selected cells
    ClipWait
    Loop, parse, Clipboard, `n, `r ; Splitting cells before putting them in an array
    {
        arr.Push(A_LoopField)
    }
}


/*
************Press Alt+P for starting the adding process*********************
*/

!p::
howmany := Array1.MaxIndex() - 1

; we need to simulate pressing the Enter with Javascript (only after adding the values of course), but be aware in your case it can be other events like "blur" etc
js1=
(
var keyboardEvent = new KeyboardEvent('keyup', {
    code: 'Enter',
    key: 'Enter',
    charCode: 13,
    keyCode: 13
})
)
page.Evaluate(js1)

Loop, % howmany
{
element := Array1[A_Index]
element2:= Chrome.Jxon_Dump(element)                     ; you pass the Autohotkey variable to Javascript
js=
(
document.querySelector('#example div.table-caption input').value = %element2%
document.querySelector('#example div.table-caption input').dispatchEvent(new Event('input'))
)
page.Evaluate(js)           ; fills a search bar with the value from the first Excel column
Sleep, 800
page.Evaluate("document.querySelector('#example div.rsp-price__text').dispatchEvent(new MouseEvent('click'))")
Sleep, 800

element := Array2[A_Index]
element2:= Chrome.Jxon_Dump(element)                      ; you pass the Autohotkey variable to Javascript
js2=
(
document.querySelector('#example table > tr:nth-child(2) input.eci-textinput').value = %element2%
document.querySelector('#example table > tr:nth-child(2) input.eci-textinput').dispatchEvent(new Event('input'))
)
page.Evaluate(js2)              ; fills a specified cell with the value from the second Excel column
Sleep, 400

; simulating pressing the Enter button
js3=
(
document.querySelector("#example table > tr:nth-child(2) input.eci-textinput").dispatchEvent(keyboardEvent)
)
page.Evaluate(js3)
Sleep, 400
}
Return
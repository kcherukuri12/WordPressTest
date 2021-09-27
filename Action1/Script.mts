
'Create an account in wordpress - 
'''steps: open browser and enter url https://wordpress.com/start/user?user_email=
'''enter enter email address; choose username and password and click create your account button

'Login to wordpress after creating an account-
'''steps: open browser and enter url https://wordpress.com/log-in/
'''enter email address or user name and click continue button
'''enter password and click login button
'''redirects to a page https://wordpress.com/read; click on the avatar icon on the top right side and it redirects to my profile page

'Open Chrome and open Wordpress Login page
sURL = "https://wordpress.com/log-in"
sBrow = "chrome.exe"
Systemutil.Run sBrow, sURL

Dim WPLogin, Obj1
Set WPLogin = Browser ("title:=.*").Page("micclass:=Page")
'Verification of opening a wordpress login page 
If WPLogin.Link("name:=Back.*").Exist(10) Then
 Reporter.ReportEvent micPass,"Test 1: Wordpress Login navigation"," Wordpress Login page navigation successful"  
 @@ script infofile_;_ZIP::ssf30.xml_;_
'Enter valid email address
Set  usernameOrEmail = description.Create()
 usernameOrEmail("micclass").value="WebEdit"
 usernameOrEmail("html tag").value = "INPUT"
 usernameOrEmail("name").value = "usernameOrEmail"
 usernameOrEmail("type").value = "text"
 usernameOrEmail("html id").value = "usernameOrEmail"
WPLogin.webedit(usernameOrEmail).Click 'Enter usernameOrEmail 
wait 2
Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "kcherukuri1@gmail.com"
wait 2

'Validation of valid email
'Set InvalidEmail = description.Create()
'InvalidEmail("micclass").value="WebElement"
'InvalidEmail("innertext").value="Log in to your accountEmail.*"
'InvalidEmail("html tag").value="DIV"
'InvalidEmail("visible").value="True"
'If Not WPLogin.WebElement(InvalidEmail).Exist(3) Then
'Reporter.ReportEvent micPass, "Valid Email Entry successful"
'End IF 
'Click continue
Set Continue = description.Create()
Continue("micclass").value="WebButton"
Continue("html tag").value = "BUTTON"
Continue("name").value = "Continue"
Continue("type").value = "submit"
WPLogin.webbutton(Continue).Click 'Click on Continue button after entering valid email under Wordpress Login page

'Enter Password
Set  password = description.Create()
password("micclass").value="WebEdit"
password("html tag").value = "INPUT"
password("name").value = "password"
password("type").value = "password"
password("html id").value = "password"
wait 2
'Password to be entered when prompted and seen as encrypted after entering
Dim InpPass, PassString
wait 4
InpPass = InputBox ("Enter Password")
PassString = Crypt.Encrypt(InpPass)
WPLogin.webedit(password).SetSecure PassString
'Click Login
Set LogIn = description.Create()
LogIn("micclass").value="WebButton"
LogIn("html tag").value = "BUTTON"
LogIn("name").value = "Log In"
LogIn("type").value = "submit"
WPLogin.webbutton(LogIn).Click 'Click on Login button

'Myprofile page
'sURL = "https://wordpress.com/me"
'sBrow = "chrome.exe"
'Systemutil.Run sBrow, sURL

'Click My Profile icon
Set MyProfile = description.Create()
MyProfile("micclass").value="Image"
MyProfile("html tag").value = "IMG"
MyProfile("url").value = "https://wordpress.com/me"
MyProfile("alt").value = "My Profile"
MyProfile("image type").value = "Image Link"
'MyProfile("src").value = "https://2.gravatar.com/avatar/bcb4b7ded3ad7d1e5a4368b786dde46b?s=96&d=mm"
'MyProfile("xpath").value = "//A/SPAN[normalize-space()=My Profile]/IMG[1]"
wait (2)
Set shellObj=CreateObject("WScript.Shell")
shellObj.SendKeys "{PGUP}"
WPLogin.image(MyProfile).Click 
wait 5
'My Profile Add/Edit First name
Set  FirstName = description.Create()
FirstName("micclass").value="WebEdit"
FirstName("html tag").value = "INPUT"
 FirstName("name").value = "first_name"
 FirstName("type").value = "text"
 FirstName("html id").value = "first_name"
'WPLogin.webedit(FirstName).Click 
WPLogin.webedit(FirstName).Set "FirstName" 

'My Profile Add/Edit Last name
Set  LastName = description.Create()
 LastName("micclass").value="WebEdit"
 LastName("html tag").value = "INPUT"
 LastName("name").value = "last_name"
 LastName("type").value = "text"
 LastName("html id").value = "last_name"
'WPLogin.webedit(LastName).Click 
WPLogin.webedit(LastName).Set "LastName"  @@ script infofile_;_ZIP::ssf44.xml_;_
 @@ script infofile_;_ZIP::ssf45.xml_;_
'My Profile Add/Edit Public display name
Set  PublicDisplayName = description.Create()
PublicDisplayName("micclass").value="WebEdit"
PublicDisplayName("html tag").value = "INPUT"
PublicDisplayName("name").value = "display_name"
PublicDisplayName("type").value = "text"
PublicDisplayName("html id").value = "display_name"
'WPLogin.webedit(PublicDisplayName).Click 
WPLogin.webedit(PublicDisplayName).Set "PublicDisplayName"

'My Profile  Add/Edit About me
Set  AboutMe = description.Create()
AboutMe("micclass").value="WebEdit"
AboutMe("html tag").value = "TEXTAREA"
AboutMe("name").value = "description"
AboutMe("type").value = "textarea"
AboutMe("html id").value = "description"

set WshShell2 = CreateObject("WScript.Shell")
WPLogin.webedit(AboutMe).Set "I am a software tester"
WPLogin.webedit(AboutMe).Click 
WshShell2.SendKeys"{ENTER}"

wait 2
'My Profile SaveProfileDetails
Set SaveProfileDetails = description.Create()
SaveProfileDetails("micclass").value="WebButton"
SaveProfileDetails("html tag").value = "BUTTON"
SaveProfileDetails("name").value = "Save profile details"
SaveProfileDetails("type").value = "submit"
Browser ("title:=.*").Page("micclass:=Page").webbutton(SaveProfileDetails).Click
'My Profile SaveProfileDetails Verification
If Browser ("title:=.*").Page("title:=.*").webbutton("name:=Dismiss","type:=submit").Exist(10) Then
 Browser ("title:=.*").Page("title:=.*").webbutton("name:=Dismiss","type:=submit").Click
 Reporter.ReportEvent micPass,"Wordpress SaveProfileDetails"," My Profile save profile details successful"  
End  If

'Add Profile Links @@ script infofile_;_ZIP::ssf45.xml_;_
If Browser ("title:=.*").Page("title:=.*").webbutton("name:=Add","type:=button","disabled:=0").Exist(5) Then
 Browser ("title:=.*").Page("title:=.*").webbutton("name:=Add","type:=button","disabled:=0").Click
 Reporter.ReportEvent micPass,"Wordpress AddProfileLinks"," My Profile add profile links successful"  
'Select Add word press site
         If Browser ("title:=.*").Page("title:=.*").WebMenu("name:=Add WordPress Site Add").Exist(3) Then
             Browser ("title:=.*").Page("title:=.*").WebMenu("name:=Add WordPress Site Add").Select (1)
'Click Cancel
               If Browser ("title:=.*").Page("title:=.*").webbutton("name:=Cancel","type:=submit").Exist(2) Then
	           Browser ("title:=.*").Page("title:=.*").webbutton("name:=Cancel","type:=submit").Click
                     End If
                       End  If
                           End  If @@ script infofile_;_ZIP::ssf47.xml_;_


 'Click Logout
 Function WordPressLogout
Set LogOut = description.Create()
LogOut("micclass").value="WebButton"
LogOut("html tag").value = "BUTTON"
LogOut("name").value = "Log out"
LogOut("type").value = "button"
Browser ("title:=.*").Page("micclass:=Page").webbutton(LogOut).Click 'Click on Logout button
End  Function @@ script infofile_;_ZIP::ssf29.xml_;_

Call WordPressLogout

Else
 Reporter.ReportEvent micFail,"Test 1: Wordpress Loginpage navigation","Please verify Wordpress Login  navigation manually"  
   End If 
   Browser("title:=.*").Close

 @@ script infofile_;_ZIP::ssf36.xml_;_
 @@ script infofile_;_ZIP::ssf38.xml_;_
 @@ script infofile_;_ZIP::ssf39.xml_;_
 @@ script infofile_;_ZIP::ssf35.xml_;_
 @@ script infofile_;_ZIP::ssf44.xml_;_
 @@ script infofile_;_ZIP::ssf55.xml_;_
 @@ script infofile_;_ZIP::ssf54.xml_;_
 @@ script infofile_;_ZIP::ssf57.xml_;_
 @@ script infofile_;_ZIP::ssf53.xml_;_

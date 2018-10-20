<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML 2 //EN">
<HTML>
<HEAD>
<style>
a:hover {color : navy;}
</style>
<TITLE>TextMe Updater</TITLE>
</HEAD>
<BODY>
<?php
if ($_GET['v'] == '1.5.22') {
	echo '<h1>No Updates Available</h1>';
	echo '<p><i>You have the latest version.<br>Please check again later, thank you.</i></p>';
	echo '<p><a href="https://github.com/pmachapman/TextMe/" target="github">Visit our website</a></p>';
} elseif ($_GET['v'] == '1.5.9999') {
	echo '<h1>No Updates Available</h1>';
	echo '<p><i>Thank you for testing TextMe 1.6 Beta.<br>Please check again later, thank you.</i></p>';
	echo '<p><a href="https://github.com/pmachapman/TextMe/" target="github">Visit our website</a></p>';
	echo '<p><a href="https://github.com/pmachapman/TextMe/issues" target="github">Submit a bug report</a></p>';
} else {
	echo '<h1>TextMe 1.5.22 Available</h1>';
	echo '<p>This update is highly recommended because it fixes many bugs.</p>';
	echo '<p>This version requires:</p>';
	echo '<ul><li>Windows 95/NT 4.0 or higher</li><li>Internet Explorer 4.0 Or Higher (Bundled with Windows 98 or Higher)</li>';
	echo '<li>An Internet Connection (For Update Checker)</li>';
	echo '<li><a href="files/vbrun60sp6.exe">Visual Basic 6.0 SP6 Runtime Files</a></li></ul></p>';
	echo '<p><a href="files/textme.zip">Download Now</a> (46KB)<br>Select Open when the dialog appears. You need a zip program such as <a href="https://www.7-zip.org/" target="7zip">7-Zip</a> to open this file.</p>';
	echo '<p><a href="https://github.com/pmachapman/TextMe/" target="github">Visit our website</a></p>';
}
?>
</BODY>
</HTML>

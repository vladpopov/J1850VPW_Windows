# http://cherrytree.at/misc/vk.htm
# Play/Pause VK_MEDIA_PLAY_PAUSE

$obj = New-Object -ComObject wscript.shell
$obj.SendKeys([char]179)
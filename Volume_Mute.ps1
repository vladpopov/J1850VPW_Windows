# http://cherrytree.at/misc/vk.htm
# Volume Mute VK_VOLUME_MUTE
function Date {
    Get-Date -Format "yyyy.MM.dd HH:mm:ss.fff"
    }

Write-Output "$(Date) Start execution"

(New-Object -ComObject wscript.shell).SendKeys([char]173)

Write-Output "$(Date) Stop execution"
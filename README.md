Prerequisites
-------------
Addin code is hosted on: https://danilkorotenko.github.io/OutlookAddin-SimpleBlock-ShowTaskPane/index.html

[Manifest](/manifest.xml)

Issue Description
-----------------
Insight message notification action button is outside user focus area and visually unnoticable in outlook on the web.


Steps To Reproduce
------------------
1. Install manifest: [Manifest](/manifest.xml)
2. Try to send any message to any address.
3. Look at notification message and its action button.

Actual Result
-------------
On Mac, notification message action button looks good and is in user focus area.
![](/screenshot_mac.png)

On Web version, the action button is somewhere far from user focus area, especially on wide screens.

![](/screenshot_web.png)


Expected result
---------------
On Web version, the action button is placed near notification message.

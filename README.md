Prerequisites
-------------
Addin code is hosted on: https://danilkorotenko.github.io/OutlookAddin-getAttachmentContent/index.html

[Manifest](/manifest.xml)

Issue Description
-----------------
Attachments content cannot be obtained.


Steps To Reproduce
------------------
1. Install manifest: [Manifest](/manifest.xml)
2. Send any message with attachment.
3. Look at addin logs.

Actual Result
-------------

Log message in logs
```
Obtain attachment error: OSF.DDA.Error {name: 'InvalidAttachmentId', message: 'The attachment ID was invalid.', code: 9003}
```

Expected result
---------------
Log message in logs
```
Obtain attachment error: OSF.DDA.Error {name: 'InvalidAttachmentId', message: 'The attachment ID was invalid.', code: 9003}
```

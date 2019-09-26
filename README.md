# Google Team Drive Orphaned File Restorer
If you restore trashed files that were in trashed folders, they can be orphaned in the root. This console application can restore them to their original folders. See [this StackOverflow question](https://stackoverflow.com/questions/58082724/google-drive-api-recovering-the-original-folder-for-a-file-restored-from-trash)

# Usage

You need a credentials.json file for both the Google Drive API v3 and Google Reports API (part of the G Suite SDK)

[.NET Quickstart for Google Drive API v3](https://developers.google.com/drive/api/v3/quickstart/dotnet)

[.NET Quickstart for G Suite SDK Reports API](https://developers.google.com/admin-sdk/reports/v1/quickstart/dotnet)

Use command line parameters such as:

`--move-start=2019-09-24 00:00:00 --move-end=2019-09-25 00:00:00 --drive-id=DRIVE-ID-HERE --ip=X.X.X.X`

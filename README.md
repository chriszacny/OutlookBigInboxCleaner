OutlookBigInboxCleaner
======================
This is a quick tool I put together that can be used if attempts to use **Delete All Items** or other such methods on large Inbox folders in Outlook fail with "Out of Memory" errors. Effectively, it acts as an automated bot that manually steps through and attempts to delete items in said Inbox in batches of your choosing.

## Requirements ##

1. Visual Studio 2012
2. .NET Framework 4.5
2. Outlook 2010
3. Visual Studio Tools for Microsoft Office (VSTO)

## Deployment / Usage ##

At the moment, there is no automated deployment of this Outlook Addin. However, it is simple if you use Visual Studio. You should just need to open the .sln file and **Rebuild Solution**. That will deploy the Addin to your installed Outlook instance. 

At that point, see the **Cleanup** tab in the Ribbon. Select a folder you wish to remove items from (be careful; it will attempt to remove everything!) and select **Cleanup Selected Folder**. You will then be prompted to make sure this is actually what you want to do. If you want to truly delete the items forever, check **Delete Permanently**. You can also select the batch size you want to delete items in. As a FYI - there is a three second delay of deletes between batches.

## More Information ##

Please see [My Blog Post](http://www.chriszacny.com/blog/outlook_big_inbox_cleaner/) for more information on this Addin.
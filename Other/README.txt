This is just a random project. It wasn't made to be a serious thing or company. So as the creator of this, I give you permission to edit, distribute, and do anything you want with this content.
-- Other notes
recommended build method is dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:PublishReadyToRun=false
but you do need to download dotnet sooo, yea. also the reason i couldnt just include the .exe file is bc github wont allow files over 25 mb.
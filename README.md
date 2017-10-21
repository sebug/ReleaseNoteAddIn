### Release Note Add-In
An add-in that allows to insert all the necessary info of a release into a Word document.

	az group create --name releaseNoteAddInGroup --location westeurope
	az storage account create --name releasenoteaddin --location westeurope --resource-group releaseNoteAddInGroup --sku Standard_LRS
	az functionapp create --name ReleaseNoteAddIn --storage-account releasenoteaddin --resource-group ReleaseNoteAddInGroup --consumption-plan-location westeurope

	az storage container create --name addinstatic
	az storage blob upload --container-name addinstatic --file clientside/index.html --name index.html --content-type "text/html"
	az storage container set-permission --name addinstatic --public-access blob
	

export class PageTemplate {
    private _content: string = null;
    private _folderUrl: string = null;
    private _imageMapper: { [key: string]: string } = null;
    private _listId: string = null;
    private _siteId: string = null;
    private _webId: string = null;

    constructor(content: string, siteId: string, webId: string, listId: string, folderUrl: string, imageMapper: { [key: string]: string }) {
        this._content = content;
        this._folderUrl = folderUrl;
        this._imageMapper = imageMapper;
        this._listId = listId;
        this._siteId = siteId;
        this._webId = webId;
    }

    // Reference to the page content
    get Content(): string { return this._content; }

    // Reference to the page images
    get Images(): { [key: string]: string } { return this._imageMapper; }

    // Update the folder information for the page
    updateFolderInfo(siteId: string, webId: string, listId: string, folderUrl: string) {
        // Set the template
        this._content = this._content
            .replace(new RegExp(this._siteId, "g"), siteId)
            .replace(new RegExp(this._webId, "g"), webId)
            .replace(new RegExp(this._listId, "g"), listId)
            .replace(new RegExp(this._folderUrl, "g"), folderUrl);
    }

    // Updates the image reference using the mapper provided
    updateImageId(fileName: string, uniqueId: string) {
        // Generate the regex for this file name
        let imageUniqueId = this._imageMapper[fileName];
        this._content = this._content.replace(new RegExp(`data-uniqueid="${imageUniqueId}"`, "g"), `data-uniqueid="${uniqueId}"`);
    }
}
import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;f8806a64-1ba5-47eb-993a-fa58b0bd1b38&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;95b79a10-da59-48f9-a2a3-cc1d0a7e9c51&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1629,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;95b79a10-da59-48f9-a2a3-cc1d0a7e9c51&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&quot;imageSource&quot;&#58;&quot;/_layouts/15/images/sleektemplateimagetile.jpg&quot;&#125;,&quot;links&quot;&#58;&#123;&#125;,&quot;customMetadata&quot;&#58;&#123;&quot;imageSource&quot;&#58;&#123;&#125;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;2,&quot;title&quot;&#58;&quot;SAT - App Permissions&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - App Permissions&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""><img data-sp-prop-name="imageSource"
                    src="/_layouts/15/images/sleektemplateimagetile.jpg" /></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;e83c9f79-417f-4daf-bf2c-29bc78dded53&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;c0a00519-5e8f-466c-9c37-f96acb5b6d9e&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <p>The App Permissions will allow you to see any app registration that has been granted the Sites.Selected
                Graph API permission for this site. Once this permission has been granted, the site collection
                administrator can manage the app registration's permissions to site(s) through this application. They
                will be able to grant read/write/full control to the app registration directly through this app. The
                Sites.Selected Graph API permission must be applied by the Azure Global Administrator prior to setting
                this up.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;e83c9f79-417f-4daf-bf2c-29bc78dded53&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;2&#125;,&quot;id&quot;&#58;&quot;2cc28298-f904-4c25-86da-4f4d0de56479&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="App Permissions" data-imagenaturalheight="523"
                data-imagenaturalwidth="1454"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAppPermissions/3249223301.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="52dce8da-c9fb-41f6-a6fc-38b0b212e8f3"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="523" data-width="1454"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The current permissions will be loaded and displayed. The user will have the ability to add, update or
                remove app permissions from the site.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;e83c9f79-417f-4daf-bf2c-29bc78dded53&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;3&#125;,&quot;id&quot;&#58;&quot;6c5b5640-e27a-4775-952a-d784ee4a8b0b&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Adding Access</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Add App Permission" data-imagenaturalheight="449"
                data-imagenaturalwidth="974"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAppPermissions/34722779.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="5f77b995-7bba-40d6-b501-466ea3fe1f1e"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="449" data-width="974"
                data-widthpercentage="82.26134234210457" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">To add access to the site, you will need
                to request an Azure application registration to be created and grated the
                <strong>Sites.Selected</strong> graph api permission. Once this is completed, the following information
                will be entered in to the form&#58;</p>
            <ul class="customListStyle" style="list-style-type&#58;disc;">
                <li data-list-item-id="e538c5528d454ec0e6293c6f1d914ef38">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">App ID - The application
                        registration id to give permissions to</p>
                </li>
                <li data-list-item-id="ebef03a4a38bf4e4df1bbcbe63ed5502b">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">App Name - The application
                        registion name to save as a reference</p>
                </li>
                <li data-list-item-id="e6740f483e2ef00cb41b7f715b72a75fc">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Permission</p>
                    <ul class="customListStyle">
                        <li data-list-item-id="e3866bd5d9eb034d88f000859f0f3cac8">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Read - Grants read to
                                all content in the site</p>
                        </li>
                        <li data-list-item-id="e5a6935662497e832cb6231fb8f4a45da">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Write - Grants
                                read/write to all content in the site</p>
                        </li>
                        <li data-list-item-id="e24b24e261df496c63d42f480230a877c">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Full Control - Grants
                                full permissions to the site</p>
                        </li>
                    </ul>
                </li>
            </ul>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;e83c9f79-417f-4daf-bf2c-29bc78dded53&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;4&#125;,&quot;id&quot;&#58;&quot;af7c2a61-7aa4-4fa0-a2d9-6a5771743593&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Modifying Access</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAppPermissions/82759635.png"
                data-uploading="0" data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="90385fee-0659-4710-a620-328014c3d272" data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219"
                data-imagenaturalheight="444" data-imagenaturalwidth="968" data-widthpercentage="81.75532383222307"
                data-width="968" data-height="444"></div>
            <p>To update access, select the permission to change to and then Update to confirm the request.</p>
            <p class="noSpacingAbove spacingBelow" aria-hidden="true" data-text-type="withSpacing">&#160;</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
export const AppPermissions = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Demo/SiteAssets/SitePages/SATAppPermissions", {
    "3249223301": "52dce8da-c9fb-41f6-a6fc-38b0b212e8f3",
    "34722779": "5f77b995-7bba-40d6-b501-466ea3fe1f1e",
    "82759635": "90385fee-0659-4710-a620-328014c3d272"
});
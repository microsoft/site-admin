import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;bd9e2e5b-43be-45cf-8bfe-01856b570269&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1469,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;4,&quot;title&quot;&#58;&quot;SAT - Search Documents&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - Search Documents&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1,&quot;imageSource&quot;&#58;&quot;&quot;,&quot;altText&quot;&#58;&quot;&quot;&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;b44aa8cb-eba1-4414-9c98-87d35bdad970&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;26f22015-2dea-4bf2-aa00-301475733953&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <h4 class="headingSpacingAbove headingSpacingBelow">Overview&#58;</h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Documents report will help
                identify sensitive content based on keyword or regex pattern(s). The output of this report will allow
                you to bulk label files based on keyword or regex pattern. The keyword search option will utilize the
                search api, so the site must be in the search index for this to work. The <strong>Search Type</strong>
                setting will determine if you will search by keyword or regex pattern.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;5d1cc03b-1f63-4ba0-a50c-134a3fd955c7&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h3 class="headingSpacingAbove headingSpacingBelow">Regex Search</h3>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Settings&#58;</h4>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReportsSearchDocs/2794747446.png"
                data-uploading="0" data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="1733e585-8c84-474d-b9f5-bcb07f070df6" data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219"
                data-imagenaturalheight="925" data-imagenaturalwidth="1218" data-widthpercentage="100" data-width="1218"
                data-height="925"></div>
            <p>&#160;</p>
            <p>The report can be run against all sites, or a single site. The “Skip Large List” setting can be used to
                quickly scan all content in the site collection. If this settings is used, then you will need to run the
                report against the skipped libraries from the “List/Library” tab. The Regex Search report is limited to
                the docx, pdf, pptx, xlsx and text files. The ability to target specific file types can be set in the
                <strong>File Types</strong> textbox. If left blank, then all files types will be searched. Regex
                patterns can be configured in the webpart which will allow the user to select pre-defined regex patterns
                by the defined category. The <strong>Regex Patterns</strong> textbox will display each pattern to use on
                a new line.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Output&#58;</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false,&quot;originalHeight&quot;&#58;0,&quot;originalWidth&quot;&#58;0&#125;"
                data-captiontext="Search Documents Regex Results" data-imagenaturalheight="729"
                data-imagenaturalwidth="1455" data-imageshapedata="&#123;&#125;"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1937731342.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="813edef9-28ba-47de-be91-eaab1af6c461"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="729" data-width="1455"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display the file information, sensitivity label, pattern matching,
                overshared flag and permissions. The user will have the ability to bulk label all files in the report or
                each one individually. The <strong>Secure File</strong> button will break inheritance on the file and
                remove the associated groups that have been deemed overshared from the file.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;4,&quot;zoneId&quot;&#58;&quot;27eb3b60-8af8-4341-9ae4-3e9f0ecf755e&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;ab5544e7-02d3-4c59-b319-c6998a4955d5&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true&#125;">
        <div data-sp-rte="">
            <h3 class="headingSpacingAbove headingSpacingBelow">Keyword Search</h3>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Settings&#58;</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false&#125;"
                data-captiontext="Search Documents Keyword Settings" data-imagenaturalheight="658"
                data-imagenaturalwidth="1455" data-imageshapedata="&#123;&#125;"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1753774094.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="b6a5bc4c-5b1e-4bc9-b6b6-2563df6b0836"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="658" data-width="1455"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The report can be run against all sites,
                or a single site. The “Skip Large List” setting can be used to quickly scan all content in the site
                collection. If this settings is used, then you will need to run the report against the skipped libraries
                from the “List/Library” tab. The keywords to search for can be entered in the <strong>Search
                    Terms</strong> textbox using a space to separate each one. You can use quotes for keywords that
                require match multiple words together. The ability to target specific file types can be set in the
                <strong>File Types</strong> textbox. If left blank, then all files types will be searched.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Output&#58;</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false,&quot;originalHeight&quot;&#58;null,&quot;originalWidth&quot;&#58;null&#125;"
                data-captiontext="Search Documents Keyword Results" data-imagenaturalheight="693"
                data-imagenaturalwidth="1451" data-imageshapedata="&#123;&#125;"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2892564598.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="9e32428e-16fa-44ff-9b22-bb54dd164261"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="693" data-width="1451"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove noSpacingBelow" data-text-type="noSpacing">The output of the report will display
                the highlighted summary from the search api response of finding for the keywords. The user will have the
                ability to view the file, set the sensitivity label on the file or delete the file.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
// Input Params: content, siteId, webId, listId, folderUrl, imageMapper
export const ReportSearchDocs = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReportsSearchDocs", {
    "2794747446": "1733e585-8c84-474d-b9f5-bcb07f070df6",
    "1937731342": "813edef9-28ba-47de-be91-eaab1af6c461",
    "1753774094": "b6a5bc4c-5b1e-4bc9-b6b6-2563df6b0836",
    "2892564598": "9e32428e-16fa-44ff-9b22-bb54dd164261"
});
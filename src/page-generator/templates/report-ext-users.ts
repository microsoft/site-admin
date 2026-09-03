import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;bd9e2e5b-43be-45cf-8bfe-01856b570269&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1469,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;4,&quot;title&quot;&#58;&quot;SAT - External Users&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - External Users&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1,&quot;imageSource&quot;&#58;&quot;&quot;,&quot;altText&quot;&#58;&quot;&quot;&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;11ab02eb-83a6-43f0-8744-52ea3bfb9eb1&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h4 class="headingSpacingAbove headingSpacingBelow">Overview&#58;</h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The External Users report will help
                identify external users with access to the site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Settings&#58;</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Users Settings" data-imagenaturalheight="260"
                data-imagenaturalwidth="1456"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3772245627.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="45741a83-0972-4d23-9354-9a93ff36b849"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="260" data-width="1456"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The report can be run against the site
                collection, or a single site.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Output&#58;</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Users Results" data-imagenaturalheight="537"
                data-imagenaturalwidth="1450"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3539902807.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="a161ed67-4892-4736-9d3b-cc8e7fdc83a7"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="537" data-width="1450"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The output of the report will allow you
                review the external user accounts with links to the associated group they are in. You will have the
                ability to remove the user from a specified group. The delete action will remove the user from all
                groups, removing access to the site.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
// Input Params: content, siteId, webId, listId, folderUrl, imageMapper
export const AppPage = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Demo/SiteAssets/SitePages/SATAuditReportsExtUsers.aspx", {
    "3772245627": "45741a83-0972-4d23-9354-9a93ff36b849",
    "3539902807": "a161ed67-4892-4736-9d3b-cc8e7fdc83a7"
});
import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;bd9e2e5b-43be-45cf-8bfe-01856b570269&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1469,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;4,&quot;title&quot;&#58;&quot;SAT - Search Property&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - Search Property&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1,&quot;imageSource&quot;&#58;&quot;&quot;,&quot;altText&quot;&#58;&quot;&quot;&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;b1f56a7b-c95e-4326-a175-2dda3c13c37b&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h4 class="headingSpacingAbove headingSpacingBelow">Overview&#58;</h4>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Property report is optional
                to use and requires configuration for it to be used. This report will allow the tenant admin to utilize
                a custom site property and the search managed property to tag sites. The idea behind this feature is to
                allow site admins/owners to tag their sites for ownership of organizations. The report will help
                identify sites that belong to organizations, as well as sites that are not tagged (Orphaned) which can
                be considered inactive and potentially archived. This report utilizes the search api, so only sites that
                are in the search index will be displayed in the report.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Settings</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Property Settings" data-imagenaturalheight="259"
                data-imagenaturalwidth="1457"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/4222049231.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="c5f4866d-29f6-43ed-ac83-7bfb7718c6ea"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="259" data-width="1457"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The dropdown selection will be pre-populated with the organization values that you can tag a site with.
                If the blank option is selected, then it will display the sites that currently have not been tagged.</p>
            <h4 class="headingSpacingAbove headingSpacingBelow">Report Output&#58;</h4>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Property Results" data-imagenaturalheight="296"
                data-imagenaturalwidth="1457"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1034045247.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="9a9d0963-cfd3-4d23-9273-5639bf671e03"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="296" data-width="1457"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The output of the report will show all
                sites by the selected tag. The admin will be able to view the metrics to see if any sites are no long
                being used. This will help with determining what sites should be considered for archiving.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Property Orphaned Sites" data-imagenaturalheight="535"
                data-imagenaturalwidth="1455"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2897641524.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="e7a4e4a1-f3b7-4ded-8633-08a3bc1e1253"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="535" data-width="1455"
                data-widthpercentage="100" data-uploading="0"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false&#125;"
                data-imageshapedata="&#123;&#125;"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">If no tag is selected, then the output
                of the report will show all sites that have not been tagged. This can help identify orphaned sites that
                do not currently have an owner.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
// Input Params: content, siteId, webId, listId, folderUrl, imageMapper
export const AppPage = new PageTemplate(CanvasContent, "",
    "", "",
    "/sites/Demo/SiteAssets/SitePages/SATAuditReportsSearchProperty.aspx", {
    "4222049231": "c5f4866d-29f6-43ed-ac83-7bfb7718c6ea",
    "1034045247": "9a9d0963-cfd3-4d23-9273-5639bf671e03",
    "2897641524": "e7a4e4a1-f3b7-4ded-8633-08a3bc1e1253"
});
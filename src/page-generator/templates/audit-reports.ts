import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;bd9e2e5b-43be-45cf-8bfe-01856b570269&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1629,&quot;reservedHeight&quot;&#58;320&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;5681b39a-f288-40e6-8868-bbb6fab5e3d2&quot;,&quot;title&quot;&#58;&quot;Title area&quot;,&quot;description&quot;&#58;&quot;Title area description&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;imageSourceType&quot;&#58;4,&quot;title&quot;&#58;&quot;SAT - Audit Reports&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showPublishDate&quot;&#58;false,&quot;showTopicHeader&quot;&#58;false,&quot;layoutType&quot;&#58;&quot;CutInShape&quot;,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;SAT - Audit Reports&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;authorByline&quot;&#58;[],&quot;encodedImage&quot;&#58;&quot;BBR&#58;HGxufQ~qj[fQ&quot;,&quot;headingLevel&quot;&#58;1,&quot;imageSource&quot;&#58;&quot;&quot;,&quot;altText&quot;&#58;&quot;&quot;&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;54fe42bc-9ec4-446f-a389-3b59f45cc95f&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;ba82666e-d7ac-4097-ac1a-ae8c0d392a21&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Audit Tools tab will display the
                available reports for auditing your site collection and sub-webs. The List/Libraries tab will also have
                an option to run reports directly against them. All reports will have the option to export the results
                to a CSV file for review with the admins/owners of the site.</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><strong>[Placeholder to add what reports
                    to run]</strong></p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;54fe42bc-9ec4-446f-a389-3b59f45cc95f&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;2&#125;,&quot;id&quot;&#58;&quot;86e72020-10f3-43a5-ade1-cc0c657f8e55&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Audit Tools" data-imagenaturalheight="686"
                data-imagenaturalwidth="1452"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1539840079.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="5f225b4c-199a-4831-bd09-c033216bacc8"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="686" data-width="1452"
                data-widthpercentage="100" data-uploading="0"></div>
            <h2 class="headingSpacingAbove headingSpacingBelow">Large Webs/Lists/Libraries</h2>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Target Web will default to search
                all sub-webs for the site. You are able to select a single target web to run the reports against to
                better handle site collections with a large number of sub-webs. Additionally, the Skip Large Lists will
                default to analyzing all lists in the site. This will allow you to run the report against all webs, then
                each large list individually.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;79bd59df-1a92-4994-be79-8b008410108f&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Data Loss Prevention</h2>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="DLP Report" data-imagenaturalheight="453"
                data-imagenaturalwidth="1458"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2299432232.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="e61178f9-571a-4005-901e-35501abfdee5"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="453" data-width="1458"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Data Loss Prevention (DLP) report
                will display any file that has a DLP policy applied to it. You have the ability to filter by file
                type(s).</p>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="DLP Report Results" data-imagenaturalheight="730"
                data-imagenaturalwidth="1459"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2978296301.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="927878bc-377d-4be3-ac0f-8e81e6a705f9"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="730" data-width="1459"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display the file information and DLP policies applied to them. A flag will
                be displayed to determine if the file has been overshared. This will be flagged if the file has been
                shared with the Everyone Except External Users (EEEU) or Everyone group. The webpart can be configured
                to include custom M365 groups that have been deemed large and therefore flag a file as overshared.</p>
            <p class="noSpacingAbove spacingBelow" aria-hidden="true" data-text-type="withSpacing">&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The first step is to review the
                permissions for the file, by clicking on the View Permissions button. The Secure File button will
                display a confirmation for removing the large groups from having access. This will break inheritance on
                the file and remove the large group's permissions from them.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;2&#125;,&quot;id&quot;&#58;&quot;d3e47b5a-1dbe-4151-9b15-b6f1112173e6&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Document Retention</h2>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Document Retention" data-imagenaturalheight="259"
                data-imagenaturalwidth="1451"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1920538907.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="f3ea3a70-f276-402f-b390-115a236787cb"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="259" data-width="1451"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Document Retention report will find
                content older than a specified date. This will help find stale content that should be flagged for review
                and archived/deleted. This report uses the search api to find the content, so the web must be in the
                search index.</p>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Document Retention Results" data-imagenaturalheight="556"
                data-imagenaturalwidth="1454"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3004706479.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="35ecac81-0cb7-42c6-a68b-2b6a529019c4"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="556" data-width="1454"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will allow the user to view, download or delete the file. It's recommended to
                export the results to a CSV for review and determination of archiving/deleting the files.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;3&#125;,&quot;id&quot;&#58;&quot;d3ae9843-6211-47dc-90fa-1c651004aa44&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">External Shares</h2>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Shares" data-imagenaturalheight="163"
                data-imagenaturalwidth="1457"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3510901274.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="0236ec2e-bd69-4d7c-952e-3f818bffa489"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="163" data-width="1457"
                data-widthpercentage="100" data-uploading="0"></div>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The External Shares report will find
                content that is currently viewable externally. This report uses the search api's ViewableByExternalUsers
                property to determine the files to display for the report, so the web must be in the search index.</p>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Shares Results" data-imagenaturalheight="490"
                data-imagenaturalwidth="1454"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3039756776.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="3e679b59-98d0-4ab1-a1ab-1142e8a97536"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="490" data-width="1454"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will allow the user to view, download or delete the file. It's recommended to
                export the results to a CSV for review and determination of external access.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;4&#125;,&quot;id&quot;&#58;&quot;11ab02eb-83a6-43f0-8744-52ea3bfb9eb1&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">External Users</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="External Users" data-imagenaturalheight="260"
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
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The External Users report will display
                all external accounts that have accessed the site from it's user information list. This will help
                determine if external users have access to your site.</p>
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
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;5&#125;,&quot;id&quot;&#58;&quot;cdaca3b5-0481-49a9-9b3d-1b85a628e421&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Permissions</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false&#125;"
                data-captiontext="Permissions" data-imagenaturalheight="260" data-imagenaturalwidth="1457"
                data-imageshapedata="&#123;&#125;"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/4007861782.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="6cb469c6-73f8-48a4-be59-4fde6a7eb6dd"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="260" data-width="1457"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Permissions report will display all role assignments associated with the site. This will allow you to
                review the permissions for the site and determine if it's currently being overshared.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Permission Results" data-imagenaturalheight="521"
                data-imagenaturalwidth="1453"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3602715702.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="5bf5dab7-59b5-4241-b4f5-255365e06f23"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="521" data-width="1453"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display flags determining if the site is overshared. The permission types
                are visible with M365 and Site Groups expanded. The user will have the ability to view the users or
                access the group directly.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;6&#125;,&quot;id&quot;&#58;&quot;1e84e752-f020-43f4-b1b6-fe78d2351693&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Search Agents</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Agents" data-imagenaturalheight="361"
                data-imagenaturalwidth="1456"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2930042662.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="73049d40-8fbc-469c-9aea-6f896ec1db91"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="361" data-width="1456"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Search Agents report will find any file with the .agent extension to determine if any Copilot Agents
                exist on this site.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;7&#125;,&quot;id&quot;&#58;&quot;5d1cc03b-1f63-4ba0-a50c-134a3fd955c7&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Search Documents</h2>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Search Documents report is meant to
                find content that should be securely labelled. The goal of this report is to find sensitive content and
                ensure they are secured and labelled.</p>
            <h3 class="headingSpacingAbove headingSpacingBelow">Keyword Search</h3>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false&#125;"
                data-captiontext="Search Documents Keyword" data-imagenaturalheight="658" data-imagenaturalwidth="1455"
                data-imageshapedata="&#123;&#125;"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1753774094.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="b6a5bc4c-5b1e-4bc9-b6b6-2563df6b0836"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="658" data-width="1455"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Keyword Search report utilizes the search api for finding content by the specified search terms. The
                search terms are separated by spaces and use quotes for terms that require multiple words.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2892564598.png"
                data-uploading="0" data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="9e32428e-16fa-44ff-9b22-bb54dd164261" data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219"
                data-imagenaturalheight="693" data-imagenaturalwidth="1451" data-widthpercentage="100" data-width="1451"
                data-height="693" data-captiontext="Search Documents Keyword Results"></div>
            <p>The output of the report will display the highlighted summary from the search api response of finding the
                keyword. The user will have the ability to view the file, set the sensitivity label on the file or
                delete the file.</p>
            <h3 class="headingSpacingAbove headingSpacingBelow">Regex Search&#160;</h3>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Documents Regex" data-imagenaturalheight="981"
                data-imagenaturalwidth="1454"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1366143702.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="569b728c-3248-4049-953a-deb45fa397ba"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="981" data-width="1454"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Regex Search report is limited to the docx, pdf, pptx, xlsx and text files.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1937731342.png"
                data-uploading="0" data-overlaystylesoverlaytransparency="0" data-overlaystylesoverlaycolor="light"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextboxcolor="dark"
                data-overlaystylesisitalic="false" data-overlaystylesisbold="false" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-uniqueid="813edef9-28ba-47de-be91-eaab1af6c461" data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219"
                data-imagenaturalheight="729" data-imagenaturalwidth="1455" data-widthpercentage="100" data-width="1455"
                data-height="729" data-captiontext="Search Documents Regex Results"></div>
            <p>The output of the report will display the file information, sensitivity label, pattern matching,
                overshared flag and permissions. The user will have the ability to bulk label all files in the report or
                each one individually.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;8&#125;,&quot;id&quot;&#58;&quot;0fb7cbf6-9915-4b1c-a6dc-db3c7841f17a&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Search EEEU</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search EEEU" data-imagenaturalheight="523"
                data-imagenaturalwidth="1453"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3672734863.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="2f0d49d0-6b6f-46ce-b9a2-96906870021d"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="523" data-width="1453"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Search Everyone Except External Users (EEEU) report will search the site for the EEEU and Everyone
                groups. If they are found, then the site, web, list, file or item will be identified and displayed for
                review. The Include Overshared Groups will also include any custom M365 Groups that are deemed
                overshared. The In Depth Search option will take longer to run, as it will inspect all list items that
                have unique permissions to determine if it's being overshared.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search EEEU Results" data-imagenaturalheight="475"
                data-imagenaturalwidth="1450"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/860781773.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="145d8cf9-0960-4c42-8572-e395eecdc85e"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="475" data-width="1450"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display the site, group, list, file or item that is overshared. The user
                will have the option to remove the permission or remove the unique permissions on an item. It is
                recommended to export the report to a CSV file and review it with the site admins and owners to review
                the permissions for the site where these groups are being used.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;9&#125;,&quot;id&quot;&#58;&quot;b1f56a7b-c95e-4326-a175-2dda3c13c37b&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Search Property</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search Property" data-imagenaturalheight="259"
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
            <p>The Search Property report is optional to use. It allows the tenant admin to utilize a custom site
                property and the search managed property to tag sites. The idea behind this feature is to allow
                organizations to tag sites to various organizations.</p>
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
                data-widthpercentage="100" data-uploading="0"></div>
            <p>&#160;</p>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">If no tag is selected, then the output
                of the report will show all sites that have not been tagged. This can help identify orphaned sites that
                do not currently have an owner.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;10&#125;,&quot;id&quot;&#58;&quot;b36458f1-1762-4902-bdf0-5406c945d665&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Search Users</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center"
                data-advancedimageeditordata="&#123;&quot;isAdvancedEdited&quot;&#58;false&#125;"
                data-captiontext="Search Users" data-imagenaturalheight="483" data-imagenaturalwidth="1450"
                data-imageshapedata="&#123;&#125;"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3795830073.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="8a6cb996-c1df-48d6-8556-773fbea6c229"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="483" data-width="1450"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Search Users report will allow you to find users by account or text. This is helpful when offboarding
                users or searching for user's permission to content in the site.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Search User Results" data-imagenaturalheight="477"
                data-imagenaturalwidth="1457"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2243122120.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="b9516fce-a771-45e4-8fb5-c257a1904e1e"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="477" data-width="1457"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display the account and associated group. The user will have the ability to
                remove the user from the site which will also remove them from all groups containing the user. The user
                will also have the ability to remove the user from groups they are a member of.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;11&#125;,&quot;id&quot;&#58;&quot;3f9f33aa-3681-494e-8517-a2144ed442e3&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Sensitivity Labels</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Sensitivity Labels" data-imagenaturalheight="560"
                data-imagenaturalwidth="1452"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1952315463.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="1812db47-266a-40fc-b15d-21c4aa5d5c42"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="560" data-width="1452"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Sensitivity Label report will allow you to search for files with or without a label. The ability to
                search for specified label(s) is also available. Similar to the DLP report, an Overshared flag will be
                associated with the results. By default the EEEU and Everyone groups will flag the file as overshared.
                Custom M365 Groups can be configured to help flag content as overshared.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Sensitivity Label Results" data-imagenaturalheight="689"
                data-imagenaturalwidth="1453"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/3517625209.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="b23a25f2-fa8c-450b-9bee-43e05fd096e5"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="689" data-width="1453"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display the file information, sensitivity label, flag determining if it's
                overshared and permissions. The user will have the ability to view the file permissions, secure the file
                if deemed overshared and set a label on the file.</p>
            <p class="noSpacingAbove spacingBelow" aria-hidden="true" data-text-type="withSpacing">&#160;</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;12&#125;,&quot;id&quot;&#58;&quot;4b72539b-42f6-4d57-8641-049c37324894&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Sharing Links</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Sharing Links" data-imagenaturalheight="160"
                data-imagenaturalwidth="1456"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1036008522.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="2c703103-d4c1-4bb5-8130-199f4f857f5b"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="160" data-width="1456"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Sharing Links report will display all hidden site groups that are associated with sharing content.
                When you share a file, the inheritance is broken and a hidden site group is created and associated with
                it.</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Sharing Links Results" data-imagenaturalheight="235"
                data-imagenaturalwidth="1454"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1392219256.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="32bb143c-b8c3-4563-b517-db10d888d3b4"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="235" data-width="1454"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will display the file and group information. The user will have the ability to
                view the file, group or delete the sharing link group.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;3,&quot;zoneId&quot;&#58;&quot;6ba4430d-4fdc-4d66-a904-22d8e78401e5&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;13&#125;,&quot;id&quot;&#58;&quot;13b170e8-7178-4887-bb7b-a7d5864b5d25&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h2 class="headingSpacingAbove headingSpacingBelow">Unique Permissions</h2>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Unique Permissions" data-imagenaturalheight="358"
                data-imagenaturalwidth="1456"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/1051374010.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="985982ea-ddaf-40af-905f-106bd9705b3a"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="358" data-width="1456"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The Unique Permissions report will find all files and items that have unique permissions.</p>
            <p>&#160;</p>
            <div class="imagePlugin" style="background-color&#58;transparent;position&#58;relative;"
                data-alignment="Center" data-captiontext="Unique Permissions Results" data-imagenaturalheight="723"
                data-imagenaturalwidth="1455"
                data-imageurl="/sites/Demo/site-admin/SiteAssets/SitePages/SATAuditReports/2928770881.png"
                data-listid="416cc6dd-2088-4197-91e2-3f6f7b4c6219" data-overlaystylesisbold="false"
                data-overlaystylesisitalic="false" data-overlaystylesoverlaycolor="light"
                data-overlaystylesoverlaytransparency="0" data-overlaystylestextboxcolor="dark"
                data-overlaystylestextboxopacity="0.54" data-overlaystylestextcolor="light"
                data-overlaytextstyles="&#123;&quot;textColor&quot;&#58;&quot;light&quot;,&quot;isBold&quot;&#58;false,&quot;isItalic&quot;&#58;false,&quot;textBoxColor&quot;&#58;&quot;dark&quot;,&quot;textBoxOpacity&quot;&#58;0.54,&quot;overlayColor&quot;&#58;&quot;light&quot;,&quot;overlayTransparency&quot;&#58;0&#125;"
                data-siteid="8a25e50f-a830-4e23-bfd3-38aaa20b57ba" data-uniqueid="97cfa8fc-5081-4198-bcff-4a59f4ba786a"
                data-webid="caf33edc-08f1-46f1-be6f-3946a595ebc3" data-height="723" data-width="1455"
                data-widthpercentage="100" data-uploading="0"></div>
            <p>The output of the report will contain the file or item information, along with the permissions. The user
                will be able to view the file/item or restore the permissions to inherit from its parent.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5,&quot;indentationVersion&quot;&#58;2&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
export const AuditReports = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Demo/SiteAssets/SitePages/SATAuditReports", {
        "1539840079": "5f225b4c-199a-4831-bd09-c033216bacc8",
        "2299432232": "e61178f9-571a-4005-901e-35501abfdee5",
        "2978296301": "927878bc-377d-4be3-ac0f-8e81e6a705f9",
        "1920538907": "f3ea3a70-f276-402f-b390-115a236787cb",
        "3004706479": "35ecac81-0cb7-42c6-a68b-2b6a529019c4",
        "3510901274": "0236ec2e-bd69-4d7c-952e-3f818bffa489",
        "3039756776": "3e679b59-98d0-4ab1-a1ab-1142e8a97536",
        "3772245627": "45741a83-0972-4d23-9354-9a93ff36b849",
        "3539902807": "a161ed67-4892-4736-9d3b-cc8e7fdc83a7",
        "4007861782": "6cb469c6-73f8-48a4-be59-4fde6a7eb6dd",
        "3602715702": "5bf5dab7-59b5-4241-b4f5-255365e06f23",
        "2930042662": "73049d40-8fbc-469c-9aea-6f896ec1db91",
        "1753774094": "b6a5bc4c-5b1e-4bc9-b6b6-2563df6b0836",
        "2892564598": "9e32428e-16fa-44ff-9b22-bb54dd164261",
        "1366143702": "569b728c-3248-4049-953a-deb45fa397ba",
        "1937731342": "813edef9-28ba-47de-be91-eaab1af6c461",
        "3672734863": "2f0d49d0-6b6f-46ce-b9a2-96906870021d",
        "860781773": "145d8cf9-0960-4c42-8572-e395eecdc85e",
        "4222049231": "c5f4866d-29f6-43ed-ac83-7bfb7718c6ea",
        "1034045247": "9a9d0963-cfd3-4d23-9273-5639bf671e03",
        "2897641524": "e7a4e4a1-f3b7-4ded-8633-08a3bc1e1253",
        "3795830073": "8a6cb996-c1df-48d6-8556-773fbea6c229",
        "2243122120": "b9516fce-a771-45e4-8fb5-c257a1904e1e",
        "1952315463": "1812db47-266a-40fc-b15d-21c4aa5d5c42",
        "3517625209": "b23a25f2-fa8c-450b-9bee-43e05fd096e5",
        "1036008522": "2c703103-d4c1-4bb5-8130-199f4f857f5b",
        "1392219256": "32bb143c-b8c3-4563-b517-db10d888d3b4",
        "1051374010": "985982ea-ddaf-40af-905f-106bd9705b3a",
        "2928770881": "97cfa8fc-5081-4198-bcff-4a59f4ba786a"
    });
import { PageTemplate } from "../template";

// The html content for the webpart
const CanvasContent = `<div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;1,&quot;zoneId&quot;&#58;&quot;fed1887c-145d-42fc-a459-ffc27f0cd3fa&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;0,&quot;controlIndex&quot;&#58;1&#125;,&quot;emphasis&quot;&#58;&#123;&quot;zoneEmphasis&quot;&#58;0&#125;,&quot;id&quot;&#58;&quot;2cb8b5b5-f6dd-47ca-8740-8acc94e95714&quot;,&quot;controlType&quot;&#58;3,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;webPartId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;reservedWidth&quot;&#58;1629,&quot;reservedHeight&quot;&#58;228&#125;">
        <div data-sp-webpart="" data-sp-webpartdataversion="1.7"
            data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;2cb8b5b5-f6dd-47ca-8740-8acc94e95714&quot;,&quot;title&quot;&#58;&quot;Banner&quot;,&quot;audiences&quot;&#58;[],&quot;hideOn&quot;&#58;&#123;&quot;mobile&quot;&#58;false&#125;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.7&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Site Admin Tool&quot;,&quot;imageSourceType&quot;&#58;4,&quot;layoutType&quot;&#58;&quot;FullWidthImage&quot;,&quot;textAlignment&quot;&#58;&quot;Left&quot;,&quot;showTopicHeader&quot;&#58;false,&quot;showPublishDate&quot;&#58;false,&quot;topicHeader&quot;&#58;&quot;&quot;,&quot;enableGradientEffect&quot;&#58;true,&quot;isDecorative&quot;&#58;true,&quot;isFullWidth&quot;&#58;true,&quot;authorByline&quot;&#58;[],&quot;authors&quot;&#58;[],&quot;customContentDropSupport&quot;&#58;&quot;externallink&quot;,&quot;htmlTitle&quot;&#58;&quot;&lt;h1&gt;Site Admin Tool&lt;/h1&gt;&quot;,&quot;showTimeToRead&quot;&#58;false,&quot;headingLevel&quot;&#58;1&#125;,&quot;containsDynamicDataSource&quot;&#58;false&#125;">
            <div data-sp-componentid="cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"></div>
            <div data-sp-htmlproperties=""></div>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;d58c21ec-30bc-46b9-952e-fcf6f8f3dd93&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;1&#125;,&quot;id&quot;&#58;&quot;460776bc-d093-4a88-9b2e-fae74c3ca017&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h1>Overview</h1>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">The Site Admin Tool (SAT) is designed to
                help site collection administrators (SCA) and site collection owners manage their sites. The tool will
                generate reports to help with auditing purposes for apply your governance and policies.</p>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;position&quot;&#58;&#123;&quot;layoutIndex&quot;&#58;1,&quot;zoneIndex&quot;&#58;2,&quot;zoneId&quot;&#58;&quot;d58c21ec-30bc-46b9-952e-fcf6f8f3dd93&quot;,&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;controlIndex&quot;&#58;2&#125;,&quot;id&quot;&#58;&quot;5ce345a6-883b-4c14-a9c3-970e2615c383&quot;,&quot;controlType&quot;&#58;4,&quot;isFromSectionTemplate&quot;&#58;false,&quot;addedFromPersistedData&quot;&#58;true,&quot;contentVersion&quot;&#58;5&#125;">
        <div data-sp-rte="">
            <h1>Data Readiness Tasks</h1>
            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Follow each task below for ensuring your
                site meets our governance and policies.</p>
            <ul class="customListStyle" style="list-style-type&#58;disc;">
                <li data-list-item-id="ea73ae2b3b656f5f262649095950b609d">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><a
                            href="/sites/demo/site-admin/sitepages/SATSiteConfiguration.aspx">Site Configuration</a></p>
                    <ul class="customListStyle">
                        <li data-list-item-id="e059fb4e82e85412c58e9169d46e42ba7">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Site Collection Admin
                                Only</p>
                        </li>
                        <li data-list-item-id="e4c4f00d06a72cd4ee4b404ee5f1a5b45">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Review site
                                configuration</p>
                        </li>
                    </ul>
                </li>
                <li data-list-item-id="edae68bac6218733731f93daf00f63bb1">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><a
                            href="/sites/demo/site-admin/sitepages/SATAuditReports.aspx">Audit Reports</a></p>
                    <ul class="customListStyle">
                        <li data-list-item-id="e46abafc5572b4235016fa2e81c70fc5e">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Site Collection Admin
                                &amp; Site Owners</p>
                        </li>
                        <li data-list-item-id="e9933f94e948be5a95ff549ff94bb3549">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Run various reports
                                against the site for review</p>
                        </li>
                    </ul>
                </li>
                <li data-list-item-id="e6321a3a8c46e931eb6cd4eb038e91cc9">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><a
                            href="/sites/demo/site-admin/sitepages/SATAppPermissions.aspx">App Permissions</a></p>
                    <ul class="customListStyle">
                        <li data-list-item-id="e46959215689b8f1b0f04131cc0e86476">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Site Collection Admin
                                Only</p>
                        </li>
                        <li data-list-item-id="eb4942f3838e7b53b704f584e7ab17a1e">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Review app registrations
                                granted access to the site</p>
                        </li>
                    </ul>
                </li>
                <li data-list-item-id="ebb53ba62bd3edb9f69d2306cf809287a">
                    <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing"><a
                            href="/sites/demo/site-admin/sitepages/SATFAQ.aspx">FAQ</a></p>
                    <ul class="customListStyle">
                        <li data-list-item-id="e4a8d1e26c356882700670655c16e3eb4">
                            <p class="noSpacingAbove spacingBelow" data-text-type="withSpacing">Review for handling
                                large lists or webs</p>
                        </li>
                    </ul>
                </li>
            </ul>
        </div>
    </div>
    <div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"
        data-sp-controldata="&#123;&quot;controlType&quot;&#58;0,&quot;pageSettingsSlice&quot;&#58;&#123;&quot;isDefaultDescription&quot;&#58;true,&quot;isAIGeneratedDescription&quot;&#58;false,&quot;isDefaultThumbnail&quot;&#58;true,&quot;isSpellCheckEnabled&quot;&#58;true,&quot;globalRichTextStylingVersion&quot;&#58;1,&quot;rtePageSettings&quot;&#58;&#123;&quot;contentVersion&quot;&#58;5&#125;,&quot;isEmailReady&quot;&#58;false,&quot;webPartsPageSettings&quot;&#58;&#123;&quot;isTitleHeadingLevelsEnabled&quot;&#58;true,&quot;isLowQualityImagePlaceholderEnabled&quot;&#58;true&#125;&#125;&#125;">
    </div>
</div>`;

// Export the page
export const SAT = new PageTemplate(CanvasContent, "8a25e50f-a830-4e23-bfd3-38aaa20b57ba",
    "caf33edc-08f1-46f1-be6f-3946a595ebc3", "416cc6dd-2088-4197-91e2-3f6f7b4c6219",
    "/sites/Dev/SiteAssets/SitePages/PageGenerator", {});
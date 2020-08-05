var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './MyFirstTeamsTabWebPart.module.scss';
import * as strings from 'MyFirstTeamsTabWebPartStrings';
var MyFirstTeamsTabWebPart = /** @class */ (function (_super) {
    __extends(MyFirstTeamsTabWebPart, _super);
    function MyFirstTeamsTabWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    MyFirstTeamsTabWebPart.prototype.render = function () {
        var title = '';
        var subTitle = '';
        var siteTabTitle = '';
        if (this.context.sdks.microsoftTeams) {
            // We have teams context for the web part
            title = "Welcome to Teams!";
            subTitle = "Building custom enterprise tabs for your business.";
            siteTabTitle = "We are in the context of following Team: " + this.context.sdks.microsoftTeams.context.teamName;
        }
        else {
            // We are rendered in normal SharePoint context
            title = "Welcome to SharePoint!";
            subTitle = "Customize SharePoint experiences using Web Parts.";
            siteTabTitle = "We are in the context of following site: " + this.context.pageContext.web.title;
        }
        this.domElement.innerHTML = "\n    <div class=\"" + styles.myFirstTeamsTab + "\">\n      <div class=\"" + styles.container + "\">\n        <div class=\"" + styles.row + "\">\n          <div class=\"" + styles.column + "\">\n            <span class=\"" + styles.title + "\">" + title + "</span>\n            <p class=\"" + styles.subTitle + "\">" + subTitle + "</p>\n            <p class=\"" + styles.description + "\">" + siteTabTitle + "</p>\n            <p class=\"" + styles.description + "\">Description property value - " + escape(this.properties.description) + "</p>\n            <a href=\"https://aka.ms/spfx\" class=\"" + styles.button + "\">\n              <span class=\"" + styles.label + "\">Learn more</span>\n            </a>\n          </div>\n        </div>\n      </div>\n    </div>";
    };
    Object.defineProperty(MyFirstTeamsTabWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    MyFirstTeamsTabWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return MyFirstTeamsTabWebPart;
}(BaseClientSideWebPart));
export default MyFirstTeamsTabWebPart;
//# sourceMappingURL=MyFirstTeamsTabWebPart.js.map
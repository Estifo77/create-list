var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from "react";
import styles from "./Hello.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import SPService from "../../../_services/SPServices";
var testFields = [
    {
        Title: "TextField",
        FieldTypeKind: 2,
    },
    {
        Title: "Number",
        FieldTypeKind: 3,
    },
    {
        Title: "Date",
        FieldTypeKind: 4,
    },
    {
        Title: "User",
        FieldTypeKind: 20,
    },
];
var Hello = (function (_super) {
    __extends(Hello, _super);
    function Hello() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.spService = new SPService(_this.props.context);
        return _this;
    }
    Hello.prototype.componentDidMount = function () {
        // this.spService.createList("SampleTestList");
        // this.spService.createSiteField("fieldone","SampleTestList")
        // this.spService.createSiteForAList("Column_one","SampleTestList")
        this.spService.createFieldsForAList("SampleTestList", testFields);
    };
    Hello.prototype.render = function () {
        return (React.createElement("div", { className: styles.hello },
            React.createElement("div", { className: styles.container },
                React.createElement("div", { className: styles.row },
                    React.createElement("div", { className: styles.column },
                        React.createElement("span", { className: styles.title }, "Welcome to SharePoint!"),
                        React.createElement("p", { className: styles.subTitle }, "Customize SharePoint experiences using Web Parts."),
                        React.createElement("p", { className: styles.description }, escape(this.props.description)),
                        React.createElement("a", { href: "https://aka.ms/spfx", className: styles.button },
                            React.createElement("span", { className: styles.label }, "Learn more")))))));
    };
    return Hello;
}(React.Component));
export default Hello;

//# sourceMappingURL=Hello.js.map

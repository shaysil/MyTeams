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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import * as strings from 'MyTeamsWebPartStrings';
import { MyTeams } from './components/myTeams';
var MyTeamsWebPart = (function (_super) {
    __extends(MyTeamsWebPart, _super);
    function MyTeamsWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._getTenantInfo = function () { return __awaiter(_this, void 0, void 0, function () {
            var tenant, tenantResponse, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        tenant = null;
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._graphClient.api('organization').select('id').version('v1.0').get()];
                    case 2:
                        tenantResponse = _a.sent();
                        tenant = tenantResponse.value;
                        console.log(tenant);
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.log('Error getting tenant information');
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/, tenant];
                }
            });
        }); };
        return _this;
    }
    MyTeamsWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _a = this;
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient()];
                    case 1:
                        _a._graphClient = _c.sent();
                        if (!(!this.properties.tenantInfo && this.properties.openInClientApp)) return [3 /*break*/, 3];
                        _b = this.properties;
                        return [4 /*yield*/, this._getTenantInfo()];
                    case 2:
                        _b.tenantInfo = _c.sent();
                        _c.label = 3;
                    case 3: return [2 /*return*/, _super.prototype.onInit.call(this)];
                }
            });
        });
    };
    MyTeamsWebPart.prototype.render = function () {
        return __awaiter(this, void 0, void 0, function () {
            var element;
            return __generator(this, function (_a) {
                element = React.createElement(MyTeams, {
                    graphClient: this._graphClient,
                    tenantId: this.properties.tenantInfo.id,
                    openInClientApp: this.properties.openInClientApp
                });
                ReactDom.render(element, this.domElement);
                return [2 /*return*/];
            });
        });
    };
    MyTeamsWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(MyTeamsWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    MyTeamsWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                PropertyPaneToggle('openInClientApp', {
                                    label: strings.OpenInClientAppFieldLabel,
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return MyTeamsWebPart;
}(BaseClientSideWebPart));
export default MyTeamsWebPart;
//# sourceMappingURL=MyTeamsWebPart.js.map
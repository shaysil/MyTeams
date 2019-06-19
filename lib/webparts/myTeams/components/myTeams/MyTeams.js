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
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { Persona } from 'office-ui-fabric-react/lib/Persona';
import { List } from 'office-ui-fabric-react/lib/List';
import styles from '../myTeams/MyTeams.module.scss';
var MyTeams = (function (_super) {
    __extends(MyTeams, _super);
    function MyTeams(props) {
        var _this = _super.call(this, props) || this;
        _this._myTeams = [];
        _this._load = function () { return __awaiter(_this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this;
                        return [4 /*yield*/, this._getTeams()];
                    case 1:
                        _a._myTeams = _b.sent();
                        this.setState({
                            items: this._myTeams
                        });
                        return [2 /*return*/];
                }
            });
        }); };
        _this._onRenderCell = function (team, index) {
            return (React.createElement("div", { className: styles.card, onClick: _this._openChannel.bind(_this, team.id, _this.props.tenantId) },
                React.createElement("div", { className: styles.containerCard },
                    React.createElement(Persona, { text: team.displayName, hidePersonaDetails: true, coinSize: 48, className: styles.initials }),
                    React.createElement("span", null, team.displayName))));
        };
        _this._openChannel = function (teamId, tenantId) { return __awaiter(_this, void 0, void 0, function () {
            var link, teamChannels, channel;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        link = '#';
                        return [4 /*yield*/, this._getTeamChannels(teamId)];
                    case 1:
                        teamChannels = _a.sent();
                        channel = teamChannels[0];
                        if (this.props.openInClientApp) {
                            link = "https://teams.microsoft.com/l/channel/" + channel.id + "/" + channel.displayName + "?groupId=" + teamId + "&tenantId=" + tenantId;
                        }
                        else {
                            link = "https://teams.microsoft.com/_#/conversations/" + channel.displayName + "?threadId=" + channel.id + "&ctx=channel";
                        }
                        window.open(link, '_blank');
                        return [2 /*return*/];
                }
            });
        }); };
        _this._getTeams = function () { return __awaiter(_this, void 0, void 0, function () {
            var myTeams, teamsResponse, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        myTeams = [];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        debugger;
                        return [4 /*yield*/, this.props.graphClient.api('me/joinedTeams').version('v1.0').get()];
                    case 2:
                        teamsResponse = _a.sent();
                        myTeams = teamsResponse.value;
                        console.log(myTeams);
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.log('Error getting teams');
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/, myTeams];
                }
            });
        }); };
        _this._getTeamChannels = function (teamId) { return __awaiter(_this, void 0, void 0, function () {
            var channels, channelsResponse, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        channels = [];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.props.graphClient.api("teams/" + teamId + "/channels").version('v1.0').get()];
                    case 2:
                        channelsResponse = _a.sent();
                        channels = channelsResponse.value;
                        return [3 /*break*/, 4];
                    case 3:
                        error_2 = _a.sent();
                        console.log('Error getting channels for team ' + teamId);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/, channels];
                }
            });
        }); };
        _this.state = {
            items: []
        };
        return _this;
    }
    MyTeams.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._load()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    MyTeams.prototype.componentDidUpdate = function (prevProps) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(this.props.openInClientApp !== prevProps.openInClientApp)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._load()];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    MyTeams.prototype.render = function () {
        return (React.createElement(FocusZone, null,
            React.createElement(List, { className: styles.myTeams, items: this._myTeams, renderedWindowsAhead: 4, onRenderCell: this._onRenderCell })));
    };
    return MyTeams;
}(React.Component));
export { MyTeams };
//# sourceMappingURL=MyTeams.js.map
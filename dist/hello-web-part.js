define("2b9a7d43-1b56-4c16-bf44-39a5bf1c3199_0.0.1", ["react","react-dom","@microsoft/sp-core-library","@microsoft/sp-webpart-base","HelloWebPartStrings","@microsoft/sp-lodash-subset","@microsoft/sp-http"], function(__WEBPACK_EXTERNAL_MODULE_0__, __WEBPACK_EXTERNAL_MODULE_2__, __WEBPACK_EXTERNAL_MODULE_3__, __WEBPACK_EXTERNAL_MODULE_4__, __WEBPACK_EXTERNAL_MODULE_5__, __WEBPACK_EXTERNAL_MODULE_13__, __WEBPACK_EXTERNAL_MODULE_15__) { return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 1);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_0__;

/***/ }),
/* 1 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_react__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_react___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0_react__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react_dom__ = __webpack_require__(2);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react_dom___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_react_dom__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__microsoft_sp_core_library__ = __webpack_require__(3);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__microsoft_sp_core_library___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__microsoft_sp_core_library__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__microsoft_sp_webpart_base__ = __webpack_require__(4);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__microsoft_sp_webpart_base___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3__microsoft_sp_webpart_base__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4_HelloWebPartStrings__ = __webpack_require__(5);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4_HelloWebPartStrings___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_4_HelloWebPartStrings__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__components_Hello__ = __webpack_require__(6);
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






var HelloWebPart = (function (_super) {
    __extends(HelloWebPart, _super);
    function HelloWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    HelloWebPart.prototype.render = function () {
        var element = __WEBPACK_IMPORTED_MODULE_0_react__["createElement"](__WEBPACK_IMPORTED_MODULE_5__components_Hello__["a" /* default */], {
            description: this.properties.description,
            context: this.context,
        });
        __WEBPACK_IMPORTED_MODULE_1_react_dom__["render"](element, this.domElement);
    };
    HelloWebPart.prototype.onDispose = function () {
        __WEBPACK_IMPORTED_MODULE_1_react_dom__["unmountComponentAtNode"](this.domElement);
    };
    Object.defineProperty(HelloWebPart.prototype, "dataVersion", {
        get: function () {
            return __WEBPACK_IMPORTED_MODULE_2__microsoft_sp_core_library__["Version"].parse("1.0");
        },
        enumerable: true,
        configurable: true
    });
    HelloWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: __WEBPACK_IMPORTED_MODULE_4_HelloWebPartStrings__["PropertyPaneDescription"],
                    },
                    groups: [
                        {
                            groupName: __WEBPACK_IMPORTED_MODULE_4_HelloWebPartStrings__["BasicGroupName"],
                            groupFields: [
                                Object(__WEBPACK_IMPORTED_MODULE_3__microsoft_sp_webpart_base__["PropertyPaneTextField"])("description", {
                                    label: __WEBPACK_IMPORTED_MODULE_4_HelloWebPartStrings__["DescriptionFieldLabel"],
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    };
    return HelloWebPart;
}(__WEBPACK_IMPORTED_MODULE_3__microsoft_sp_webpart_base__["BaseClientSideWebPart"]));
/* harmony default export */ __webpack_exports__["default"] = (HelloWebPart);



/***/ }),
/* 2 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_2__;

/***/ }),
/* 3 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_3__;

/***/ }),
/* 4 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_4__;

/***/ }),
/* 5 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_5__;

/***/ }),
/* 6 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_react__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_react___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0_react__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__ = __webpack_require__(7);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__microsoft_sp_lodash_subset__ = __webpack_require__(13);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__microsoft_sp_lodash_subset___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__microsoft_sp_lodash_subset__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__services_SPServices__ = __webpack_require__(14);
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
        _this.spService = new __WEBPACK_IMPORTED_MODULE_3__services_SPServices__["a" /* default */](_this.props.context);
        return _this;
    }
    Hello.prototype.componentDidMount = function () {
        // this.spService.createList("SampleTestList");
        // this.spService.createSiteField("fieldone","SampleTestList")
        // this.spService.createSiteForAList("Column_one","SampleTestList")
        this.spService.createFieldsForAList("SampleTestList", testFields);
    };
    Hello.prototype.render = function () {
        return (__WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("div", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].hello },
            __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("div", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].container },
                __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("div", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].row },
                    __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("div", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].column },
                        __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("span", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].title }, "Welcome to SharePoint!"),
                        __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("p", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].subTitle }, "Customize SharePoint experiences using Web Parts."),
                        __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("p", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].description }, Object(__WEBPACK_IMPORTED_MODULE_2__microsoft_sp_lodash_subset__["escape"])(this.props.description)),
                        __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("a", { href: "https://aka.ms/spfx", className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].button },
                            __WEBPACK_IMPORTED_MODULE_0_react__["createElement"]("span", { className: __WEBPACK_IMPORTED_MODULE_1__Hello_module_scss__["a" /* default */].label }, "Learn more")))))));
    };
    return Hello;
}(__WEBPACK_IMPORTED_MODULE_0_react__["Component"]));
/* harmony default export */ __webpack_exports__["a"] = (Hello);



/***/ }),
/* 7 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* tslint:disable */
__webpack_require__(8);
var styles = {
    hello: 'hello_efddbf0d',
    container: 'container_efddbf0d',
    row: 'row_efddbf0d',
    column: 'column_efddbf0d',
    'ms-Grid': 'ms-Grid_efddbf0d',
    title: 'title_efddbf0d',
    subTitle: 'subTitle_efddbf0d',
    description: 'description_efddbf0d',
    button: 'button_efddbf0d',
    label: 'label_efddbf0d',
};
/* harmony default export */ __webpack_exports__["a"] = (styles);
/* tslint:enable */ 



/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

var content = __webpack_require__(9);
var loader = __webpack_require__(11);

if(typeof content === "string") content = [[module.i, content]];

// add the styles to the DOM
for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1], true);

if(content.locals) module.exports = content.locals;

/***/ }),
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__(10)(false);
// imports


// module
exports.push([module.i, ".hello_efddbf0d .container_efddbf0d{max-width:700px;margin:0 auto;box-shadow:0 2px 4px 0 rgba(0,0,0,.2),0 25px 50px 0 rgba(0,0,0,.1)}.hello_efddbf0d .row_efddbf0d{margin:0 -8px;box-sizing:border-box;color:\"[theme:white, default: #ffffff]\";background-color:\"[theme:themeDark, default: #005a9e]\";padding:20px}.hello_efddbf0d .row_efddbf0d:after,.hello_efddbf0d .row_efddbf0d:before{display:table;content:\"\";line-height:0}.hello_efddbf0d .row_efddbf0d:after{clear:both}.hello_efddbf0d .column_efddbf0d{position:relative;min-height:1px;padding-left:8px;padding-right:8px;box-sizing:border-box}[dir=ltr] .hello_efddbf0d .column_efddbf0d{float:left}[dir=rtl] .hello_efddbf0d .column_efddbf0d{float:right}.hello_efddbf0d .column_efddbf0d .ms-Grid_efddbf0d{padding:0}@media (min-width:640px){.hello_efddbf0d .column_efddbf0d{width:83.33333333333334%}}@media (min-width:1024px){.hello_efddbf0d .column_efddbf0d{width:66.66666666666666%}}@media (min-width:1024px){[dir=ltr] .hello_efddbf0d .column_efddbf0d{left:16.66667%}[dir=rtl] .hello_efddbf0d .column_efddbf0d{right:16.66667%}}@media (min-width:640px){[dir=ltr] .hello_efddbf0d .column_efddbf0d{left:8.33333%}[dir=rtl] .hello_efddbf0d .column_efddbf0d{right:8.33333%}}.hello_efddbf0d .title_efddbf0d{font-size:21px;font-weight:100;color:\"[theme:white, default: #ffffff]\"}.hello_efddbf0d .description_efddbf0d,.hello_efddbf0d .subTitle_efddbf0d{font-size:17px;font-weight:300;color:\"[theme:white, default: #ffffff]\"}.hello_efddbf0d .button_efddbf0d{text-decoration:none;height:32px;min-width:80px;background-color:\"[theme:themePrimary, default: #0078d7]\";border-color:\"[theme:themePrimary, default: #0078d7]\";color:\"[theme:white, default: #ffffff]\";outline:transparent;position:relative;font-family:Segoe UI WestEuropean,Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif;-webkit-font-smoothing:antialiased;font-size:14px;font-weight:400;border-width:0;text-align:center;cursor:pointer;display:inline-block;padding:0 16px}.hello_efddbf0d .button_efddbf0d .label_efddbf0d{font-weight:600;font-size:14px;height:32px;line-height:32px;margin:0 4px;vertical-align:top;display:inline-block}", ""]);

// exports


/***/ }),
/* 10 */
/***/ (function(module, exports) {

/*
	MIT License http://www.opensource.org/licenses/mit-license.php
	Author Tobias Koppers @sokra
*/
// css base code, injected by the css-loader
module.exports = function(useSourceMap) {
	var list = [];

	// return the list of modules as css string
	list.toString = function toString() {
		return this.map(function (item) {
			var content = cssWithMappingToString(item, useSourceMap);
			if(item[2]) {
				return "@media " + item[2] + "{" + content + "}";
			} else {
				return content;
			}
		}).join("");
	};

	// import a list of modules into the list
	list.i = function(modules, mediaQuery) {
		if(typeof modules === "string")
			modules = [[null, modules, ""]];
		var alreadyImportedModules = {};
		for(var i = 0; i < this.length; i++) {
			var id = this[i][0];
			if(typeof id === "number")
				alreadyImportedModules[id] = true;
		}
		for(i = 0; i < modules.length; i++) {
			var item = modules[i];
			// skip already imported module
			// this implementation is not 100% perfect for weird media query combinations
			//  when a module is imported multiple times with different media queries.
			//  I hope this will never occur (Hey this way we have smaller bundles)
			if(typeof item[0] !== "number" || !alreadyImportedModules[item[0]]) {
				if(mediaQuery && !item[2]) {
					item[2] = mediaQuery;
				} else if(mediaQuery) {
					item[2] = "(" + item[2] + ") and (" + mediaQuery + ")";
				}
				list.push(item);
			}
		}
	};
	return list;
};

function cssWithMappingToString(item, useSourceMap) {
	var content = item[1] || '';
	var cssMapping = item[3];
	if (!cssMapping) {
		return content;
	}

	if (useSourceMap && typeof btoa === 'function') {
		var sourceMapping = toComment(cssMapping);
		var sourceURLs = cssMapping.sources.map(function (source) {
			return '/*# sourceURL=' + cssMapping.sourceRoot + source + ' */'
		});

		return [content].concat(sourceURLs).concat([sourceMapping]).join('\n');
	}

	return [content].join('\n');
}

// Adapted from convert-source-map (MIT)
function toComment(sourceMap) {
	// eslint-disable-next-line no-undef
	var base64 = btoa(unescape(encodeURIComponent(JSON.stringify(sourceMap))));
	var data = 'sourceMappingURL=data:application/json;charset=utf-8;base64,' + base64;

	return '/*# ' + data + ' */';
}


/***/ }),
/* 11 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/* WEBPACK VAR INJECTION */(function(global) {
/**
 * An IThemingInstruction can specify a rawString to be preserved or a theme slot and a default value
 * to use if that slot is not specified by the theme.
 */
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
Object.defineProperty(exports, "__esModule", { value: true });
// IE needs to inject styles using cssText. However, we need to evaluate this lazily, so this
// value will initialize as undefined, and later will be set once on first loadStyles injection.
var _injectStylesWithCssText;
// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
// load-themed-styles hosted on the page.
var _root = (typeof window === 'undefined') ? global : window; // tslint:disable-line:no-any
var _themeState = initializeThemeState();
/**
 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
 */
// tslint:disable-next-line:max-line-length
var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
/** Maximum style text length, for supporting IE style restrictions. */
var MAX_STYLE_CONTENT_SIZE = 10000;
var now = function () { return (typeof performance !== 'undefined' && !!performance.now) ? performance.now() : Date.now(); };
function measure(func) {
    var start = now();
    func();
    var end = now();
    _themeState.perf.duration += end - start;
}
/**
 * initialize global state object
 */
function initializeThemeState() {
    var state = _root.__themeState__ || {
        theme: undefined,
        lastStyleElement: undefined,
        registeredStyles: []
    };
    if (!state.runState) {
        state = __assign({}, (state), { perf: {
                count: 0,
                duration: 0
            }, runState: {
                flushTimer: 0,
                mode: 0 /* sync */,
                buffer: []
            } });
    }
    if (!state.registeredThemableStyles) {
        state = __assign({}, (state), { registeredThemableStyles: [] });
    }
    _root.__themeState__ = state;
    return state;
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load
 * event is fired.
 * @param {string | ThemableArray} styles Themable style text to register.
 * @param {boolean} loadAsync When true, always load styles in async mode, irrespective of current sync mode.
 */
function loadStyles(styles, loadAsync) {
    if (loadAsync === void 0) { loadAsync = false; }
    measure(function () {
        var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
        if (_injectStylesWithCssText === undefined) {
            _injectStylesWithCssText = shouldUseCssText();
        }
        var _a = _themeState.runState, mode = _a.mode, buffer = _a.buffer, flushTimer = _a.flushTimer;
        if (loadAsync || mode === 1 /* async */) {
            buffer.push(styleParts);
            if (!flushTimer) {
                _themeState.runState.flushTimer = asyncLoadStyles();
            }
        }
        else {
            applyThemableStyles(styleParts);
        }
    });
}
exports.loadStyles = loadStyles;
/**
 * Allows for customizable loadStyles logic. e.g. for server side rendering application
 * @param {(processedStyles: string, rawStyles?: string | ThemableArray) => void}
 * a loadStyles callback that gets called when styles are loaded or reloaded
 */
function configureLoadStyles(loadStylesFn) {
    _themeState.loadStyles = loadStylesFn;
}
exports.configureLoadStyles = configureLoadStyles;
/**
 * Configure run mode of load-themable-styles
 * @param mode load-themable-styles run mode, async or sync
 */
function configureRunMode(mode) {
    _themeState.runState.mode = mode;
}
exports.configureRunMode = configureRunMode;
/**
 * external code can call flush to synchronously force processing of currently buffered styles
 */
function flush() {
    measure(function () {
        var styleArrays = _themeState.runState.buffer.slice();
        _themeState.runState.buffer = [];
        var mergedStyleArray = [].concat.apply([], styleArrays);
        if (mergedStyleArray.length > 0) {
            applyThemableStyles(mergedStyleArray);
        }
    });
}
exports.flush = flush;
/**
 * register async loadStyles
 */
function asyncLoadStyles() {
    return setTimeout(function () {
        _themeState.runState.flushTimer = 0;
        flush();
    }, 0);
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
 * is fired.
 * @param {string} styleText Style to register.
 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
 */
function applyThemableStyles(stylesArray, styleRecord) {
    if (_themeState.loadStyles) {
        _themeState.loadStyles(resolveThemableArray(stylesArray).styleString, stylesArray);
    }
    else {
        _injectStylesWithCssText ?
            registerStylesIE(stylesArray, styleRecord) :
            registerStyles(stylesArray);
    }
}
/**
 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
 * replaced.
 * @param {theme} theme JSON object of theme tokens to values.
 */
function loadTheme(theme) {
    _themeState.theme = theme;
    // reload styles.
    reloadStyles();
}
exports.loadTheme = loadTheme;
/**
 * Clear already registered style elements and style records in theme_State object
 * @option: specify which group of registered styles should be cleared.
 * Default to be both themable and non-themable styles will be cleared
 */
function clearStyles(option) {
    if (option === void 0) { option = 3 /* all */; }
    if (option === 3 /* all */ || option === 2 /* onlyNonThemable */) {
        clearStylesInternal(_themeState.registeredStyles);
        _themeState.registeredStyles = [];
    }
    if (option === 3 /* all */ || option === 1 /* onlyThemable */) {
        clearStylesInternal(_themeState.registeredThemableStyles);
        _themeState.registeredThemableStyles = [];
    }
}
exports.clearStyles = clearStyles;
function clearStylesInternal(records) {
    records.forEach(function (styleRecord) {
        var styleElement = styleRecord && styleRecord.styleElement;
        if (styleElement && styleElement.parentElement) {
            styleElement.parentElement.removeChild(styleElement);
        }
    });
}
/**
 * Reloads styles.
 */
function reloadStyles() {
    if (_themeState.theme) {
        var themableStyles = [];
        for (var _i = 0, _a = _themeState.registeredThemableStyles; _i < _a.length; _i++) {
            var styleRecord = _a[_i];
            themableStyles.push(styleRecord.themableStyle);
        }
        if (themableStyles.length > 0) {
            clearStyles(1 /* onlyThemable */);
            applyThemableStyles([].concat.apply([], themableStyles));
        }
    }
}
/**
 * Find theme tokens and replaces them with provided theme values.
 * @param {string} styles Tokenized styles to fix.
 */
function detokenize(styles) {
    if (styles) {
        styles = resolveThemableArray(splitStyles(styles)).styleString;
    }
    return styles;
}
exports.detokenize = detokenize;
/**
 * Resolves ThemingInstruction objects in an array and joins the result into a string.
 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
 */
function resolveThemableArray(splitStyleArray) {
    var theme = _themeState.theme;
    var themable = false;
    // Resolve the array of theming instructions to an array of strings.
    // Then join the array to produce the final CSS string.
    var resolvedArray = (splitStyleArray || []).map(function (currentValue) {
        var themeSlot = currentValue.theme;
        if (themeSlot) {
            themable = true;
            // A theming annotation. Resolve it.
            var themedValue = theme ? theme[themeSlot] : undefined;
            var defaultValue = currentValue.defaultValue || 'inherit';
            // Warn to console if we hit an unthemed value even when themes are provided, but only if "DEBUG" is true.
            // Allow the themedValue to be undefined to explicitly request the default value.
            if (theme && !themedValue && console && !(themeSlot in theme) && "boolean" !== 'undefined' && true) {
                console.warn("Theming value not provided for \"" + themeSlot + "\". Falling back to \"" + defaultValue + "\".");
            }
            return themedValue || defaultValue;
        }
        else {
            // A non-themable string. Preserve it.
            return currentValue.rawString;
        }
    });
    return {
        styleString: resolvedArray.join(''),
        themable: themable
    };
}
/**
 * Split tokenized CSS into an array of strings and theme specification objects
 * @param {string} styles Tokenized styles to split.
 */
function splitStyles(styles) {
    var result = [];
    if (styles) {
        var pos = 0; // Current position in styles.
        var tokenMatch = void 0; // tslint:disable-line:no-null-keyword
        while (tokenMatch = _themeTokenRegex.exec(styles)) {
            var matchIndex = tokenMatch.index;
            if (matchIndex > pos) {
                result.push({
                    rawString: styles.substring(pos, matchIndex)
                });
            }
            result.push({
                theme: tokenMatch[1],
                defaultValue: tokenMatch[2] // May be undefined
            });
            // index of the first character after the current match
            pos = _themeTokenRegex.lastIndex;
        }
        // Push the rest of the string after the last match.
        result.push({
            rawString: styles.substring(pos)
        });
    }
    return result;
}
exports.splitStyles = splitStyles;
/**
 * Registers a set of style text. If it is registered too early, we will register it when the
 * window.load event is fired.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStyles(styleArray) {
    var head = document.getElementsByTagName('head')[0];
    var styleElement = document.createElement('style');
    var _a = resolveThemableArray(styleArray), styleString = _a.styleString, themable = _a.themable;
    styleElement.type = 'text/css';
    styleElement.appendChild(document.createTextNode(styleString));
    _themeState.perf.count++;
    head.appendChild(styleElement);
    var record = {
        styleElement: styleElement,
        themableStyle: styleArray
    };
    if (themable) {
        _themeState.registeredThemableStyles.push(record);
    }
    else {
        _themeState.registeredStyles.push(record);
    }
}
/**
 * Registers a set of style text, for IE 9 and below, which has a ~30 style element limit so we need
 * to register slightly differently.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStylesIE(styleArray, styleRecord) {
    var head = document.getElementsByTagName('head')[0];
    var registeredStyles = _themeState.registeredStyles;
    var lastStyleElement = _themeState.lastStyleElement;
    var stylesheet = lastStyleElement ? lastStyleElement.styleSheet : undefined;
    var lastStyleContent = stylesheet ? stylesheet.cssText : '';
    var lastRegisteredStyle = registeredStyles[registeredStyles.length - 1];
    var resolvedStyleText = resolveThemableArray(styleArray).styleString;
    if (!lastStyleElement || (lastStyleContent.length + resolvedStyleText.length) > MAX_STYLE_CONTENT_SIZE) {
        lastStyleElement = document.createElement('style');
        lastStyleElement.type = 'text/css';
        if (styleRecord) {
            head.replaceChild(lastStyleElement, styleRecord.styleElement);
            styleRecord.styleElement = lastStyleElement;
        }
        else {
            head.appendChild(lastStyleElement);
        }
        if (!styleRecord) {
            lastRegisteredStyle = {
                styleElement: lastStyleElement,
                themableStyle: styleArray
            };
            registeredStyles.push(lastRegisteredStyle);
        }
    }
    lastStyleElement.styleSheet.cssText += detokenize(resolvedStyleText);
    Array.prototype.push.apply(lastRegisteredStyle.themableStyle, styleArray); // concat in-place
    // Preserve the theme state.
    _themeState.lastStyleElement = lastStyleElement;
}
/**
 * Checks to see if styleSheet exists as a property off of a style element.
 * This will determine if style registration should be done via cssText (<= IE9) or not
 */
function shouldUseCssText() {
    var useCSSText = false;
    if (typeof document !== 'undefined') {
        var emptyStyle = document.createElement('style');
        emptyStyle.type = 'text/css';
        useCSSText = !!emptyStyle.styleSheet;
    }
    return useCSSText;
}


/* WEBPACK VAR INJECTION */}.call(exports, __webpack_require__(12)))

/***/ }),
/* 12 */
/***/ (function(module, exports) {

var g;

// This works in non-strict mode
g = (function() {
	return this;
})();

try {
	// This works if eval is allowed (see CSP)
	g = g || Function("return this")() || (1,eval)("this");
} catch(e) {
	// This works if the window reference is available
	if(typeof window === "object")
		g = window;
}

// g can still be undefined, but nothing to do about it...
// We return undefined, instead of nothing here, so it's
// easier to handle this case. if(!global) { ...}

module.exports = g;


/***/ }),
/* 13 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_13__;

/***/ }),
/* 14 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__ = __webpack_require__(15);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__);
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

// import reject = Promise.reject;
var getHeader = {
    headers: {
        accept: "application/json;",
    },
};
var postHeader = {
    headers: {
        "content-type": "application/json;odata.metadata=full",
        accept: "application/json;odata.metadata=full",
    },
};
var deleteHeader = {
    headers: {
        "content-type": "application/json;odata.metadata=full",
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
    },
};
var updateHeader = {
    headers: {
        "content-type": "application/json;odata.metadata=full",
        accept: "application/json;odata.metadata=full",
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
    },
};
var SPService = (function () {
    function SPService(context) {
        var _this = this;
        this.context = context;
        this.webUrl = this.context.pageContext.web.absoluteUrl;
        this.serverUrl = this.context.pageContext.web.serverRelativeUrl;
        this.siteUrl = this.context.pageContext.site.absoluteUrl;
        this.loggedUserName = this.context.pageContext.user.displayName;
        this.loggedUserEmail = this.context.pageContext.user.email;
        this.loggedUserId = this.context.pageContext.legacyPageContext.userId;
        this.adWebUrl = window.location.origin + ":2023/ADExplorer";
        this.getServiceUrl = function (url) { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2 /*return*/, this.context.spHttpClient
                        .get(url, __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__["SPHttpClient"].configurations.v1, {
                        headers: getHeader.headers,
                    })
                        .then(function (response) { return __awaiter(_this, void 0, void 0, function () {
                        var jsonResponse, responseValue, jsonResponse, error, _a;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    if (!response.ok) return [3 /*break*/, 2];
                                    return [4 /*yield*/, response.json()];
                                case 1:
                                    jsonResponse = _b.sent();
                                    responseValue = {
                                        hasError: false,
                                        value: jsonResponse.value,
                                    };
                                    return [2 /*return*/, responseValue];
                                case 2: return [4 /*yield*/, response.json()];
                                case 3:
                                    jsonResponse = _b.sent();
                                    _a = {
                                        hasError: true
                                    };
                                    return [4 /*yield*/, jsonResponse.error];
                                case 4:
                                    error = (_a.error = _b.sent(),
                                        _a);
                                    return [2 /*return*/, Promise.reject(error)];
                            }
                        });
                    }); })
                        .catch(function (error) {
                        //console.error(error ? error.message : "");
                        console.error(error);
                        return error;
                    })];
            });
        }); };
        this.getDepInfo = function (depName) {
            var queryParams = {
                queryParams: {
                    AllowEmailAddresses: true,
                    AllowMultipleEntities: false,
                    AllUrlZones: false,
                    MaximumEntitySuggestions: 5,
                    PrincipalSource: 15,
                    PrincipalType: 12,
                    QueryString: depName,
                },
            };
            var url = _this.webUrl +
                "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser";
            var options = {
                headers: postHeader.headers,
                body: queryParams,
            };
            return _this.post(url, options)
                .then(function (result) {
                var resultKey = JSON.parse(result.value.value);
                var ensuruserParam;
                if (resultKey.length > 0) {
                    ensuruserParam = { logonName: resultKey[0].Key };
                }
                return _this.getFinalDepInfo(ensuruserParam);
            })
                .catch(function (err) {
                throw new Error(err);
            });
        };
    }
    SPService.prototype.getListItems = function (listName) {
        return this.context.spHttpClient
            .get(this.webUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?$top=15", __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__["SPHttpClient"].configurations.v1)
            .then(function (response) { return response.json(); })
            .then(function (data) { return data; }, function (error) { return error; });
    };
    SPService.prototype.changeDateFormat = function (date) {
        var insertedDate = new Date(date);
        var insertedDate2 = this.getFormattedResult(insertedDate.getMonth() + 1) +
            "/" +
            this.getFormattedResult(insertedDate.getDate()) +
            "/" +
            this.getFormattedResult(insertedDate.getFullYear());
        var returnedDate = insertedDate2.split("/");
        return returnedDate[2] + "-" + returnedDate[0] + "-" + returnedDate[1];
    };
    SPService.prototype.getByUrl = function (url) {
        return this.get(url, false)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.getAllItems = function (listName) {
        var url = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/items";
        return this.get(url).then(function (result) {
            return result;
        });
    };
    SPService.prototype.getFormattedResult = function (num) {
        if (num <= 9) {
            return "0" + num;
        }
        return num;
    };
    SPService.prototype.getItemById = function (listName, id) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items?$select=*,EncodedAbsUrl,FileLeafRef&$filter=Id eq " +
            id;
        // return this.get(url, false)
        return this.getServiceUrl(url)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.getFilteredItems = function (listName, query) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items" +
            query;
        return this.get(url, false)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.getFieldsChoices = function (listName, fieldName) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/fields/getByTitle('" +
            fieldName +
            "')/Choices";
        return this.get(url, false)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.editAndGet = function (listName, id, inputs) {
        var _this = this;
        return this.updateItem(listName, inputs, id).then(function (response) {
            return _this.getItemById(listName, id).then(function (json) {
                return json;
            });
        });
    };
    SPService.prototype.getAllFiles = function (serverRelativeUrl) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/files";
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getFilteredFiles = function (serverRelativeUrl, query) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/files" + query;
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getAllFolders = function (serverRelativeUrl) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/folders";
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getFilteredFolder = function (serverRelativeUrl, query) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/folders" + query;
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getLibraryInformationByName = function (libraryName) {
        var restUrl = this.webUrl + "/_api/web/folders?$filter=Name eq '" + libraryName + "'";
        return this.get(restUrl)
            .then(function (json) {
            return json;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.getInformationUsingServerRelativeUrl = function (serverRelativeUrl) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')";
        return this.get(restUrl)
            .then(function (json) {
            return json;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.getItemCount = function (listName) {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/ItemCount";
        return this.get(restUrl)
            .then(function (json) {
            return json;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.get = function (url, check) {
        var _this = this;
        if (check === void 0) { check = false; }
        return this.context.spHttpClient
            .get(url, __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__["SPHttpClient"].configurations.v1, {
            headers: getHeader.headers,
        })
            .then(function (response) { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, response.json().then(function (json) {
                        return json;
                    })];
            });
        }); })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.postItem = function (listName, data) {
        var url = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/items";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.createList = function (listName, description) {
        if (description === void 0) { description = ""; }
        var url = this.webUrl + "/_api/web/lists";
        var options = {
            headers: postHeader.headers,
            body: {
                // '__metadata': { 'type': 'SP.List' },
                BaseTemplate: 100,
                Title: listName,
                Description: description,
            },
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.createSiteField = function (fieldName, groupName) {
        var url = this.webUrl + "/_api/web/fields";
        var options = {
            headers: postHeader.headers,
            body: {
                // '__metadata': { 'type': 'SP.Field' },
                Title: fieldName,
                FieldTypeKind: 2,
                Group: groupName,
            },
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.createFieldsForAList = function (listName, fieldsDefinition) {
        var _this = this;
        Promise.all(fieldsDefinition.map(function (fieldDefinition) {
            _this.createFieldForAList(listName, fieldDefinition);
        })).then(function () {
            return;
        });
    };
    SPService.prototype.createFieldForAList = function (listName, fieldDefinition) {
        var url = this.webUrl + ("/_api/web/lists/getByTitle('" + listName + "')/fields");
        var options = {
            headers: postHeader.headers,
            body: fieldDefinition,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.postNotification = function (listName, data) {
        var url = window.location.origin + "/sites/portal/" +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.updateItem = function (listName, data, id, toJson) {
        if (toJson === void 0) { toJson = true; }
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items(" +
            id +
            ")";
        var options = {
            headers: updateHeader.headers,
            body: data,
        };
        return this.post(url, options, toJson)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.updateFileMetaData = function (fileServerRelativeUrl, data) {
        var url = this.webUrl +
            ("/_api/web/getFileByServerRelativeUrl('" + fileServerRelativeUrl + "')/ListItemAllFields");
        var options = {
            headers: updateHeader.headers,
            body: data,
        };
        return this.post(url, options, true)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.updateFolderMetaData = function (folderServerRelativeUrl, data) {
        var url = this.webUrl +
            ("/_api/web/getFolderByServerRelativeUrl('" + folderServerRelativeUrl + "')/ListItemAllFields");
        var options = {
            headers: updateHeader.headers,
            body: data,
        };
        return this.post(url, options, true)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.postFile = function (listName, file) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/RootFolder/files/add(url='" +
            file.name +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
            body: file,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    // uploadFile(serverRelativeUrl, file): Promise<any> {
    //     sp.web.getFolderByServerRelativeUrl( this.webUrl +
    //       "/_api/web/getFolderByServerRelativeUrl('" +
    //       serverRelativeUrl).files.add(file.name, file, true).then(f => {
    //
    //       f.file.getItem().then(item => {
    //         item.update({
    //           Title: "Metadata Updated"
    //         }).then((result) => {
    //
    //
    //
    //             return result;
    //
    //         }) .catch((err) => {
    //
    //           Promise.reject(err);
    //           return err;
    //         });
    //       }).catch((err) => {
    //
    //         Promise.reject(err);
    //         return err;
    //       });
    //     }).catch((err) => {
    //
    //       Promise.reject(err);
    //       return err;
    //     });
    //   }
    SPService.prototype.postFileByServerRelativeUrl = function (serverRelativeUrl, file) {
        var url = this.webUrl +
            "/_api/web/getFolderByServerRelativeUrl('" +
            serverRelativeUrl +
            "')/files/add(url='" +
            file.name +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
            body: file,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.createFile = function (listName, fileName) {
        var url = this.webUrl +
            "/_api/web/GetFolderByServerRelativeUrl('" +
            this.serverUrl +
            "/" +
            listName +
            "')/files/add(url='" +
            fileName +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.createFileByServerRelativeUrl = function (folderServerRelativeUrl, fileName) {
        var url = this.webUrl +
            "/_api/web/GetFolderByServerRelativeUrl('" +
            folderServerRelativeUrl +
            "')/files/add(url='" +
            fileName +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.createFolder = function (serverRelativeUrl, folderName) {
        var url = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/folders/add(url='" + folderName + "')";
        var options = {
            headers: postHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.moveFile = function (listName, originalFileName, newFileName) {
        var url = this.webUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            originalFileName +
            "')/moveto(newurl = '" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            newFileName +
            "', flags = 1)";
        var options = {
            headers: updateHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.moveFolder = function (listName, originalFileName, newFileName) {
        var url = this.webUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            originalFileName +
            "')/moveto(newurl = '" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            newFileName +
            "', flags = 1)";
        var options = {
            headers: updateHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.deleteItem = function (listName, id) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items(" +
            id +
            ")";
        var options = {
            headers: deleteHeader.headers,
        };
        return this.context.spHttpClient
            .post(url, __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__["SPHttpClient"].configurations.v1, options)
            .then(function (response) {
            return response.json();
        })
            .catch(function (err) {
            return err;
        });
    };
    // async post(url: string, postInformation, check = false): Promise<any> {
    //   if (check) {
    //     return await this.dateConverter
    //       .toEuropean(postInformation.body)
    //       .then((response) => {
    //         const options: ISPHttpClientOptions = {
    //           headers: postInformation.headers,
    //           body: JSON.stringify(response),
    //         };
    //         return this.context.spHttpClient
    //           .post(url, SPHttpClient.configurations.v1, options)
    //           .then((result) => {
    //             return result.json().then((json) => {
    //               return json;
    //             });
    //           })
    //           .catch((err) => {
    //
    //             Promise.reject(err);
    //           });
    //       });
    //   }
    SPService.prototype.post = function (url, postInformation, toJson) {
        if (toJson === void 0) { toJson = true; }
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var options;
            return __generator(this, function (_a) {
                options = {
                    headers: postInformation.headers,
                    body: JSON.stringify(postInformation.body),
                };
                return [2 /*return*/, this.context.spHttpClient
                        .post(url, __WEBPACK_IMPORTED_MODULE_0__microsoft_sp_http__["SPHttpClient"].configurations.v1, options)
                        .then(function (response) { return __awaiter(_this, void 0, void 0, function () {
                        var jsonResponse, responseValue, jsonResponse, error, _a;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    if (!toJson) return [3 /*break*/, 6];
                                    if (!response.ok) return [3 /*break*/, 2];
                                    return [4 /*yield*/, response.json()];
                                case 1:
                                    jsonResponse = _b.sent();
                                    responseValue = {
                                        hasError: false,
                                        value: jsonResponse,
                                    };
                                    return [2 /*return*/, responseValue];
                                case 2: return [4 /*yield*/, response.json()];
                                case 3:
                                    jsonResponse = _b.sent();
                                    _a = {
                                        hasError: true
                                    };
                                    return [4 /*yield*/, jsonResponse.error];
                                case 4:
                                    error = (_a.error = _b.sent(),
                                        _a);
                                    return [2 /*return*/, Promise.reject(error)];
                                case 5: return [3 /*break*/, 7];
                                case 6: return [2 /*return*/, response];
                                case 7: return [2 /*return*/];
                            }
                        });
                    }); })
                        .catch(function (error) {
                        //console.error(error ? error.message : "");
                        console.error(error);
                        return error;
                    })];
            });
        });
    };
    SPService.prototype.isCurrentUserInGroup = function (groupName) {
        return __awaiter(this, void 0, void 0, function () {
            var url;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = this.context.pageContext.web.absoluteUrl +
                            "/_api/Web/SiteGroups/GetByName('" +
                            groupName +
                            "')/Users?$filter=email eq '" +
                            this.loggedUserEmail +
                            "'";
                        return [4 /*yield*/, this.get(url).then(function (response) {
                                var result = false;
                                if (response.length > 0) {
                                    result = true;
                                }
                                else {
                                    result = false;
                                }
                                return result;
                            })];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    SPService.prototype.getListInformationByName = function (listName) {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')";
        return this.get(restUrl).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getAllGroupsOfAUser = function () {
        var restUrl = this.context.pageContext.web.absoluteUrl + "/_api/web/currentuser/?$expand=groups";
        return this.get(restUrl)
            .then(function (json) {
            return json.Groups;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getDepartmentsFromAD = function (Ou) {
        var depdata = [];
        return fetch(window.location.origin + ":2023/adexplorer/getorgstr?ou=" + Ou)
            .then(function (response) { return response.json(); })
            .then(function (data) {
            depdata = data;
            return depdata;
        })
            .catch(function (error) {
            console.error(error);
        });
    };
    SPService.prototype.getMyProperties = function () {
        var url = this.webUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
        return this.get(url, false).then(function (res) {
            return res;
        });
    };
    SPService.prototype.getUserDepartmentFromAD = function (userName) {
        return fetch(window.location.origin + ":2023/ADExplorer/getUserOU/?UserName=" + userName)
            .then(function (data) {
            var depdata = data;
            return depdata;
        })
            .catch(function (error) {
            console.error(error);
        });
    };
    SPService.prototype.getUserSubDepartments = function (userName, siteName) {
        return fetch(this.adWebUrl + "/GetSubOU/?OU=" + siteName + "&Parent=" + userName)
            .then(function (data) {
            var depdata = data;
            return depdata;
        })
            .catch(function (error) {
            console.error(error);
        });
    };
    SPService.prototype.getPermissionIds = function () {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('Permissions')/items?&$select=*&$orderby=Created%20desc";
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getUserRole = function (userID) {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('UserRole')/items?&$select=*&$orderby=Created%20desc&$filter=UserId eq '" + userID + "'";
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getUserRoleResources = function (roleIds) {
        var values = roleIds;
        var filterConditions = values
            .map(function (value) { return "Role/Id eq '" + value + "'"; })
            .join(" or ");
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('RoleResource')/items?&$select=*,Role/Id,PageResource/PageCode&$expand=Role,PageResource&$orderby=Created%20desc&$filter=" + filterConditions;
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getUserRolePermissions = function (roleresourceIds) {
        var values = roleresourceIds;
        var filterConditions = values
            .map(function (value) { return "RoleResourceId eq '" + value + "'"; })
            .join(" or ");
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('RolePermission')/items?$top=200&$select=*,Permission/Id&$expand=Permission&$orderby=Created%20desc&$filter=" + filterConditions;
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getPageCodes = function () {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('PageResource')/items?&$select=*&$orderby=Created%20desc";
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getParentSiteDetail = function () {
        var url = this.webUrl + "/_api/site/RootWeb";
        return this.get(url)
            .then(function (response) {
            return response;
        })
            .catch(function (err) {
            throw new Error("error");
        });
    };
    SPService.prototype.getFinalDepInfo = function (data) {
        var url = this.webUrl + "/_api/web/ensureuser";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw new Error(err);
        });
    };
    SPService.prototype.getUserDepartment = function () {
        var url = this.webUrl + "/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName='Department')?@v='" + this.context.pageContext.user.loginName + "'";
        return this.get(url)
            .then(function (response) {
            return response.value;
        })
            .catch(function (err) {
            throw new Error("error");
        });
    };
    SPService.prototype.createNotification = function (data) {
        var url = this.siteUrl +
            "/_api/web/lists/getByTitle('Notification_associated_task')/items";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    return SPService;
}());
/* harmony default export */ __webpack_exports__["a"] = (SPService);



/***/ }),
/* 15 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_15__;

/***/ })
/******/ ])});;
//# sourceMappingURL=hello-web-part.js.map
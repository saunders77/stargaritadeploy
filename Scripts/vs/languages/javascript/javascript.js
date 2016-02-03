/*---------------------------------------------------------
 * Copyright (C) Microsoft Corporation. All rights reserved.
 *--------------------------------------------------------*/
"use strict";define("vs/languages/typescript/features/tokenization",["require","exports","vs/base/strings","vs/base/arrays","vs/base/collections","vs/editor/modes/modes","vs/editor/modes/supports","vs/languages/vsxml/vsxml","vs/editor/modes/stream","vs/languages/typescript/lib/typescriptServices","vs/editor/modes/modesExtensions"],function(e,t,r,n,o,i,s,a,l,c,u){function d(e,t){var r=c.createClassifier({log:function(){}}),n=t===m.TypeScript?b:y,o=t===m.TypeScript?w:_;return{shouldGenerateEmbeddedModels:!1,getInitialState:function(){return new g(e,null,t,f.Unknown)},tokenize:function(e,t,i,s){return h(n,o,r,t,e,i,s)}}}function h(e,t,o,a,l,u){function d(e,t,r){(0===h.tokens.length||void 0!==r||n.tail(h.tokens).type!==t)&&h.tokens.push(new s.Token(e,t,r||i.Bracket.None))}void 0===u&&(u=0);var h={tokens:[],actualStopOffset:u+l.length,endState:new g(a.getMode(),a.getStateData(),a.language,f.Unknown),modeTransitions:[{startIndex:u,mode:a.getMode()}]};if(p(a,u,l,d))return h;var b=o.getClassificationsForLine(l,f.toEndOfLineState(a.errorState),!0),w=a.language===m.TypeScript,y=0;return h.endState.errorState=f.toErrorState(b.finalLexState,l,a.errorState),b.entries.forEach(function(n){var o,s=i.Bracket.None;if(n.classification===c.TokenClass.Punctuation){var h=l.charCodeAt(y);s=v[h]||i.Bracket.None,o=e[h]||t[n.classification],d(y+u,o,s)}else n.classification===c.TokenClass.Comment?a.errorState===f.InDocMultiLineCommentTrivia||l.match(/\/\*\*.*$/)&&!l.match(/^\s*\/\//)?d(y+u,w?"comment.doc.ts":"comment.doc.js",i.Bracket.None):a.errorState===f.InMultiLineCommentTrivia||l.match(/\/\*.*$/)&&!l.match(/^\s*\/\//)?d(y+u,w?"comment.block.ts":"comment.block.js",i.Bracket.None):d(y+u,w?"comment.line.ts":"comment.line.js",i.Bracket.None):d(y+u,t[n.classification]||r.empty,void 0);y+=n.length}),h}function p(e,t,r,n){if(!r.match(C))return!1;var o,s,c=new l.LineStream(r),u=new a.VSXMLExpression(e.getMode(),e);for(c.advanceIfRegExp(C).length>0&&n(t,"comment.vs",i.Bracket.None);!c.eos();){s=c.pos();do{if(o=u.tokenize(c),null===o||void 0===o||(void 0===o.type||null===o.type)&&(void 0===o.nextState||null===o.nextState))throw new Error("Tokenizer must return a valid state");if(o.nextState&&(o.nextState.setStateData(u.getStateData()),u=o.nextState),c.pos()<=s)throw new Error}while(!o.type&&""!==o.type);n(s+t,o.type,o.bracket)}return!0}!function(e){e[e.TypeScript=0]="TypeScript",e[e.EcmaScript5=1]="EcmaScript5"}(t.Language||(t.Language={}));var m=t.Language;t.createTokenizationSupport=d;var f;!function(e){e[e.Unknown=0]="Unknown",e[e.InMultiLineCommentTrivia=1]="InMultiLineCommentTrivia",e[e.InDocMultiLineCommentTrivia=2]="InDocMultiLineCommentTrivia",e[e.InSingleQuoteStringLiteral=3]="InSingleQuoteStringLiteral",e[e.InDoubleQuoteStringLiteral=4]="InDoubleQuoteStringLiteral"}(f||(f={}));var f;!function(e){function t(t){switch(t){case e.InSingleQuoteStringLiteral:return c.EndOfLineState.InSingleQuoteStringLiteral;case e.InDoubleQuoteStringLiteral:return c.EndOfLineState.InDoubleQuoteStringLiteral;case e.InMultiLineCommentTrivia:case e.InDocMultiLineCommentTrivia:return c.EndOfLineState.InMultiLineCommentTrivia}return c.EndOfLineState.Start}function r(t,r,n){switch(t){case c.EndOfLineState.Start:return e.Unknown;case c.EndOfLineState.InSingleQuoteStringLiteral:return e.InSingleQuoteStringLiteral;case c.EndOfLineState.InDoubleQuoteStringLiteral:return e.InDoubleQuoteStringLiteral;case c.EndOfLineState.InMultiLineCommentTrivia:return n===e.InDocMultiLineCommentTrivia||r.match(/\/\*\*.*$/)?e.InDocMultiLineCommentTrivia:e.InMultiLineCommentTrivia}}e.toEndOfLineState=t,e.toErrorState=r}(f||(f={}));var g=function(){function e(e,t,r,n){this._mode=e,this._state=t,this.language=r,this.errorState=n}return e.prototype.clone=function(){return new e(this._mode,u.safeStateClone(this._state),this.language,this.errorState)},e.prototype.equals=function(t){return t===this?!0:t&&t instanceof e?this.errorState!==t.errorState?!1:u.safeStateEquals(this._state,t._state):!1},e.prototype.getMode=function(){return this._mode},e.prototype.tokenize=function(){throw new Error},e.prototype.getStateData=function(){return this._state},e.prototype.setStateData=function(e){this._state=e},e}(),v=o.createNumberDictionary();v["(".charCodeAt(0)]=i.Bracket.Open,v[")".charCodeAt(0)]=i.Bracket.Close,v["{".charCodeAt(0)]=i.Bracket.Open,v["}".charCodeAt(0)]=i.Bracket.Close,v["[".charCodeAt(0)]=i.Bracket.Open,v["]".charCodeAt(0)]=i.Bracket.Close;var b=o.createNumberDictionary();b["(".charCodeAt(0)]="delimiter.parenthesis.ts",b[")".charCodeAt(0)]="delimiter.parenthesis.ts",b["{".charCodeAt(0)]="delimiter.bracket.ts",b["}".charCodeAt(0)]="delimiter.bracket.ts",b["[".charCodeAt(0)]="delimiter.array.ts",b["]".charCodeAt(0)]="delimiter.array.ts";var w=o.createNumberDictionary();w[c.TokenClass.Identifier]="identifier.ts",w[c.TokenClass.Keyword]="keyword.ts",w[c.TokenClass.Operator]="delimiter.ts",w[c.TokenClass.Punctuation]="delimiter.ts",w[c.TokenClass.NumberLiteral]="number.ts",w[c.TokenClass.RegExpLiteral]="regexp.ts",w[c.TokenClass.StringLiteral]="string.ts";var y=o.createNumberDictionary();y["(".charCodeAt(0)]="delimiter.parenthesis.js",y[")".charCodeAt(0)]="delimiter.parenthesis.js",y["{".charCodeAt(0)]="delimiter.bracket.js",y["}".charCodeAt(0)]="delimiter.bracket.js",y["[".charCodeAt(0)]="delimiter.array.js",y["]".charCodeAt(0)]="delimiter.array.js";var _=o.createNumberDictionary();_[c.TokenClass.Identifier]="identifier.js",_[c.TokenClass.Keyword]="keyword.js",_[c.TokenClass.Operator]="delimiter.js",_[c.TokenClass.Punctuation]="delimiter.js",_[c.TokenClass.NumberLiteral]="number.js",_[c.TokenClass.RegExpLiteral]="regexp.js",_[c.TokenClass.StringLiteral]="string.js";var C=/^\s*\/\/\//}),define("vs/base/lifecycle",["require","exports"],function(e,t){function r(e){for(var t=0,r=e.length;r>t;t++)e[t]&&e[t].dispose();return[]}function n(){for(var e=[],t=0;t<arguments.length;t++)e[t-0]=arguments[t];return{dispose:function(){return r(e)}}}function o(e){return{dispose:function(){return r(e)}}}function i(e){return{dispose:function(){return e()}}}function s(){for(var e=[],t=0;t<arguments.length;t++)e[t-0]=arguments[t];return o(e.map(i))}function a(e){if(e){if("function"==typeof e)return e(),null;if(Array.isArray(e)){for(;e.length>0;)e.pop()();return e}return null}return null}t.disposeAll=r,t.combinedDispose=n,t.combinedDispose2=o,t.fnToDisposable=i,t.toDisposable=s,t.cAll=a;var l=function(){function e(e){this._prefix=e,this._lastId=0}return e.prototype.generate=function(){return this._prefix+ ++this._lastId},e}();t.IdGenerator=l});var __extends=this.__extends||function(e,t){function r(){this.constructor=e}for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n]);r.prototype=t.prototype,e.prototype=new r};define("vs/base/uuid",["require","exports"],function(e,t){function r(e,t){return"function"==typeof e[t]&&0===e[t].length}function n(e){return e instanceof a||e instanceof l||r(e,"asHex",0)&&r(e,"equals",1)}function o(){return new l}function i(e){if(!c.test(e))throw new Error("invalid uuid");return new a(e)}function s(){return o().asHex()}t.isUUID=n;var a=function(){function e(e){this._value=e}return e.prototype.asHex=function(){return this._value},e.prototype.equals=function(e){return this.asHex()===e.asHex()},e}(),l=function(e){function t(){e.call(this,[t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),"-",t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),"-","4",t._randomHex(),t._randomHex(),t._randomHex(),"-",t._oneOf(t._timeHighBits),t._randomHex(),t._randomHex(),t._randomHex(),"-",t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex(),t._randomHex()].join(""))}return __extends(t,e),t._oneOf=function(e){var t=Math.floor(e.length*Math.random());return e[t]},t._randomHex=function(){return t._oneOf(t._chars)},t._chars=["0","1","2","3","4","5","6","6","7","8","9","a","b","c","d","e","f"],t._timeHighBits=["8","9","a","b"],t}(a);t.empty=new a("00000000-0000-0000-0000-000000000000"),t.v4=o;var c=/^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$/;t.parse=i,t.generateUuid=s}),define("vs/languages/typescript/features/quickFixMainActions",["require","exports","vs/base/strings","vs/base/errors","vs/base/lib/winjs.base","vs/base/paths"],function(e,t,r,n,o,i){var s=function(){function e(e){this._requestService=e.requestService,this._fileService=e.fileService,this._resourceService=e.resourceService}return e.prototype.evaluate=function(e,t,r){var n=JSON.parse(r);switch(n.type){case"addtripleslash":return this.evaluateTripleSlashQuickFix(n.name,e)}return o.Promise.as(null)},e.prototype.evaluateTripleSlashQuickFix=function(e,t){var o=this,s="/typings/"+e,a=r.rtrim(i.dirname(this._requestService.getPath("root",t)),"/"),l=i.relative(a,s),c=r.format('/// <reference path="{0}"/>\n',l),u={editOperations:[{range:{startLineNumber:1,startColumn:1,endLineNumber:1,endColumn:1},text:c}]};return this._fileService.resolveFileStat(s).then(function(){return u},function(){var t="https://raw.githubusercontent.com/borisyankov/DefinitelyTyped/master/"+e;return o._requestService.makeRequest({url:t,followRedirects:5}).then(function(e){return o._fileService.createFile(s,e.responseText).then(function(){return u},function(){return u})},function(e){return n.onUnexpectedError(r.format("Unable to fetch d.ts file at {0}: {1}",t,e.toString())),u})})},e}();t.QuickFixMainActions=s}),define("vs/languages/typescript/typescript",["require","exports","vs/platform/platform"],function(e,t,r){var n;!function(e){e.ResourceSetChanged="typescript.resourceSetChanged"}(n=t.Events||(t.Events={}));var o;!function(e){function t(e){o=e}function n(){return o}e.Identifier="typescript",r.Registry.add(e.Identifier,e);var o;e.setProjectResolver=t,e.getProjectResolver=n}(o=t.Extensions||(t.Extensions={}))});var __extends=this.__extends||function(e,t){function r(){this.constructor=e}for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n]);r.prototype=t.prototype,e.prototype=new r};define("vs/languages/typescript/typescriptMode",["require","exports","vs/base/lib/winjs.base","vs/base/lifecycle","vs/base/errors","vs/editor/modes/modesExtensions","vs/editor/modes/supports","vs/base/uuid","vs/base/arrays","vs/platform/services","vs/languages/typescript/features/tokenization","vs/languages/typescript/features/quickFixMainActions","vs/platform/thread/attribute","vs/languages/typescript/typescript"],function(e,t,r,n,o,i,s,a,l,c,u,d,h,p){var m=function(e){function t(t,r){var n=this;e.call(this,t,r,c.AsyncDescriptor.create("vs/languages/typescript/typescriptWorker2","TypeScriptWorker2")),this._disposables=[],this._projectResolverPromise=null,this._projectPromise=null,this._pendingCallbacks=[],this._telemetryService=t.telemetryService,t.eventService&&this._disposables.push(t.eventService.addListener2(p.Events.ResourceSetChanged,function(){return n._joinProjectResolver()})),this.tokenizationSupport=u.createTokenizationSupport(this,u.Language.TypeScript),this.electricCharacterSupport=new s.BracketElectricCharacterSupport(this,{brackets:[{tokenType:"delimiter.bracket.ts",open:"{",close:"}",isElectric:!0},{tokenType:"delimiter.array.ts",open:"[",close:"]",isElectric:!0},{tokenType:"delimiter.parenthesis.ts",open:"(",close:")",isElectric:!0}],docComment:{scope:"comment.doc",open:"/**",lineStart:" * ",close:" */"}}),this.referenceSupport=new s.ReferenceSupport(this,{tokens:["identifier.ts"],findReferences:function(e,t,r){return n.findReferences(e,t,r)}}),this.navigateTypesSupport=this,this.extraInfoSupport=this,this.formattingSupport=this,this.declarationSupport=new s.DeclarationSupport(this,{tokens:["identifier.ts","string.ts","attribute.value.vs"],findDeclaration:function(e,t){return n.findDeclaration(e,t)}}),this.quickFixSupport=this,this.logicalSelectionSupport=this,this.parameterHintsSupport=new s.ParameterHintsSupport(this,{triggerCharacters:["(",","],excludeTokens:["string.ts"],getParameterHints:function(e,t){return n.getParameterHints(e,t)}}),this.outlineSupport=this,this.emitOutputSupport=this,this.codeLensSupport=this,this.renameSupport=this,this.characterPairSupport=new s.CharacterPairSupport(this,{autoClosingPairs:[{open:"{",close:"}"},{open:"[",close:"]"},{open:"(",close:")"},{open:'"',close:'"',notIn:["string"]},{open:"'",close:"'",notIn:["string","comment"]},{open:"`",close:"`"}]}),this.suggestSupport=new s.SuggestSupport(this,{triggerCharacters:["."],excludeTokens:["string","comment","number"],sortBy:[{type:"reference",partSeparator:"/"}],suggest:function(e,t){return n.suggest(e,t)},getSuggestionDetails:function(e,t,r){return n.getSuggestionDetails(e,t,r)}})}return __extends(t,e),t.prototype.dispose=function(){this._disposables=n.disposeAll(this._disposables)},t.prototype._releaseCallbacks=function(e,t){for(;this._pendingCallbacks.length>0;){var r=this._pendingCallbacks.shift();t?r.error(t):r.completion(e)}},t.prototype._joinProjectResolver=function(){var e=this;if(!this._threadService.isInMainThread)return null;var t,n;if(n=new r.TPromise(function(r,n){t={completion:r,error:n},e._pendingCallbacks.push(t)},function(){var r=e._pendingCallbacks.indexOf(t);r>=0&&(e._pendingCallbacks.splice(r,1),t.error(o.canceled()))}),!this._projectPromise){if(!this._projectResolverPromise){var i=this._getProjectResolverDescription();if(!i)return null;this._projectResolverPromise=this._instantiationService.createInstance(i)}this._projectPromise=this._projectResolverPromise.then(function(e){var t=e.resolve();return r.Promise.is(t)?t:r.TPromise.as({})}).then(function(t){return t?t.hasChanges!==!0&&l.isFalsyOrEmpty(t.added)&&l.isFalsyOrEmpty(t.removed)?null:e._acceptProjectDelta(t):null});var s,c=Date.now(),u=a.v4().asHex();s=setInterval(function(){e._telemetryService.log("typescript.projectResolver.slow",{module:e._getProjectResolverDescription().moduleName,ctor:e._getProjectResolverDescription().ctorName,resolveCallUuid:u,totalWaitTime:Date.now()-c})},2e4),this._projectPromise.done(function(t){clearTimeout(s),e._projectPromise=null,e._releaseCallbacks(t,null)},function(t){clearTimeout(s),e._projectPromise=null,e._releaseCallbacks(null,t)})}return n},t.prototype._getProjectResolverDescription=function(){return p.Extensions.getProjectResolver()},t.prototype._acceptProjectDelta=function(e){return this._worker(function(t){return t.acceptProjectDelta(e)})},t.prototype.getNonWordTokenTypes=function(){return["delimiter.ts","delimiter.parenthesis.ts","delimiter.bracket.ts","delimiter.array.ts"]},t.prototype.getCommentsConfiguration=function(){return{lineCommentTokens:["//"],blockCommentStartToken:"/*",blockCommentEndToken:"*/"}},t.prototype._pickAWorkerToValidate=function(){return this._worker(function(e){return e.enableValidator()})},t.prototype.getOutline=function(e){return this._worker(function(t){return t.getOutline(e)})},t.prototype.findOccurrences=function(e,t,r){return void 0===r&&(r=!1),this._worker(function(n){return n.findOccurrences(e,t,r)})},t.prototype.suggest=function(e,t){return this._worker(function(r){return r.suggest(e,t)})},t.prototype.getSuggestionDetails=function(e,t,r){return this._worker(function(n){return n.getSuggestionDetails(e,t,r)})},t.prototype.getParameterHints=function(e,t){return this._worker(function(r){return r.getParameterHints(e,t)})},t.prototype.getEmitOutput=function(e,t){return void 0===t&&(t=void 0),this._worker(function(r){return r.getEmitOutput(e,t)})},t.prototype.getWordDefinition=function(){return t.WORD_DEFINITION},t.prototype._worker=function(t){return e.prototype._worker.call(this,t)},t.prototype.findReferences=function(e,t,r){return this._worker(function(n){return n.findReferences(e,t,r)})},Object.defineProperty(t.prototype,"filter",{get:function(){return["identifier.ts","string.ts"]},enumerable:!0,configurable:!0}),t.prototype.rename=function(e,t,r){return this._worker(function(n){return n.rename(e,t,r)})},t.prototype.getNavigateToItems=function(e){return this._worker(function(t){return t.getNavigateToItems(e)})},t.prototype.runQuickFixAction=function(e,t,r){var n=this,o=this._instantiationService.createInstance(d.QuickFixMainActions);return o.evaluate(e,t,r).then(function(o){return o?o:n.runQuickFixActionInWorker(e,t,r)})},t.prototype.runQuickFixActionInWorker=function(e,t,r){return this._worker(function(n){return n.runQuickFixAction(e,t,r)})},t.prototype.getQuickFixes=function(e,t){return this._worker(function(r){return r.getQuickFixes(e,t)})},t.prototype.getRangesToPosition=function(e,t){return this._worker(function(r){return r.getRangesToPosition(e,t)})},t.prototype.findDeclaration=function(e,t){return this._worker(function(r){return r.findDeclaration(e,t)})},t.prototype.findTypeDeclaration=function(e,t){return this._worker(function(r){return r.findTypeDeclaration(e,t)})},t.prototype.computeInfo=function(e,t){return this._worker(function(r){return r.computeInfo(e,t)})},t.prototype.getAutoFormatTriggerCharacters=function(){return[";","}","\n"]},t.prototype.formatDocument=function(e,t){return this._worker(function(r){return r.formatDocument(e,t)})},t.prototype.formatRange=function(e,t,r){return this._worker(function(n){return n.formatRange(e,t,r)})},t.prototype.formatAfterKeystroke=function(e,t,r,n){return this._worker(function(o){return o.formatAfterKeystroke(e,t,r,n)})},t.prototype.enableCodeLens=function(){return this._options&&this._options.referenceInfos===!0},t.prototype.findCodeLensSymbols=function(e){return this._worker(function(t){return t.findCodeLensSymbols(e)})},t.prototype.findCodeLensReferences=function(e,t){return this._worker(function(r){return r.findCodeLensReferences(e,t)})},t.$_acceptProjectDelta=h.AllWorkers(t,t.prototype._acceptProjectDelta),t.$_pickAWorkerToValidate=h.OneWorker(t,t.prototype._pickAWorkerToValidate,t.prototype._joinProjectResolver,c.ThreadAffinity.Group3),t.$getOutline=h.OneWorker(t,t.prototype.getOutline,c.ThreadAffinity.Group1),t.$findOccurrences=h.OneWorker(t,t.prototype.findOccurrences,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$suggest=h.OneWorker(t,t.prototype.suggest,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$getSuggestionDetails=h.OneWorker(t,t.prototype.getSuggestionDetails,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$getParameterHints=h.OneWorker(t,t.prototype.getParameterHints,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$getEmitOutput=h.OneWorker(t,t.prototype.getEmitOutput,t.prototype._joinProjectResolver,c.ThreadAffinity.Group3),t.WORD_DEFINITION=i.createWordRegExp("$"),t.$findReferences=h.OneWorker(t,t.prototype.findReferences,t.prototype._joinProjectResolver,c.ThreadAffinity.Group3),t.$rename=h.OneWorker(t,t.prototype.rename,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$getNavigateToItems=h.OneWorker(t,t.prototype.getNavigateToItems,t.prototype._joinProjectResolver,c.ThreadAffinity.Group3),t.$runQuickFixActionInWorker=h.OneWorker(t,t.prototype.runQuickFixActionInWorker,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$getQuickFixes=h.OneWorker(t,t.prototype.getQuickFixes,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$getRangesToPosition=h.OneWorker(t,t.prototype.getRangesToPosition,c.ThreadAffinity.Group1),t.$findDeclaration=h.OneWorker(t,t.prototype.findDeclaration,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$findTypeDeclaration=h.OneWorker(t,t.prototype.findTypeDeclaration,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$computeInfo=h.OneWorker(t,t.prototype.computeInfo,t.prototype._joinProjectResolver,c.ThreadAffinity.Group2),t.$formatDocument=h.OneWorker(t,t.prototype.formatDocument,c.ThreadAffinity.Group1),t.$formatRange=h.OneWorker(t,t.prototype.formatRange,c.ThreadAffinity.Group1),t.$formatAfterKeystroke=h.OneWorker(t,t.prototype.formatAfterKeystroke,c.ThreadAffinity.Group1),t.findCodeLensSymbols=h.OneWorker(t,t.prototype.findCodeLensSymbols,c.ThreadAffinity.Group1),t.findCodeLensReferences=h.OneWorker(t,t.prototype.findCodeLensReferences,t.prototype._joinProjectResolver,c.ThreadAffinity.Group3),t}(i.AbstractMode);t.TypeScriptMode=m}),define("vs/languages/javascript/javascript.extensions",["require","exports","vs/platform/platform"],function(e,t,r){var n;!function(e){function t(e){o=e}function n(){return o}e.Identifier="javascript",r.Registry.add(e.Identifier,e);var o;e.setProjectResolver=t,e.getProjectResolver=n}(n=t.Extensions||(t.Extensions={}))});var __extends=this.__extends||function(e,t){function r(){this.constructor=e}for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n]);r.prototype=t.prototype,e.prototype=new r};define("vs/languages/javascript/javascript",["require","exports","vs/base/lib/winjs.base","vs/editor/modes/modesExtensions","vs/platform/services","vs/languages/typescript/features/tokenization","vs/languages/typescript/typescriptMode","vs/editor/modes/supports","vs/languages/javascript/javascript.extensions"],function(e,t,r,n,o,i,s,a,l){var c=function(t){function s(e,r){var n=this;t.call(this,e,r),this._workerDescriptor=o.AsyncDescriptor.create("vs/languages/javascript/javascriptWorker","JavaScriptWorker"),this.referenceSupport=new a.ReferenceSupport(this,{tokens:["identifier.js"],findReferences:function(e,t,r){return n.findReferences(e,t,r)}}),this.declarationSupport=new a.DeclarationSupport(this,{tokens:["identifier.js","string.js","attribute.value.vs"],findDeclaration:function(e,t){return n.findDeclaration(e,t)}}),this.typeDeclarationSupport=void 0,this.emitOutputSupport=void 0,this.parameterHintsSupport=new a.ParameterHintsSupport(this,{triggerCharacters:["(",","],excludeTokens:["string.js","string.escape.js"],getParameterHints:function(e,t){return n.getParameterHints(e,t)}}),this.tokenizationSupport=i.createTokenizationSupport(this,i.Language.EcmaScript5),this.electricCharacterSupport=new a.BracketElectricCharacterSupport(this,{brackets:[{tokenType:"delimiter.bracket.js",open:"{",close:"}",isElectric:!0},{tokenType:"delimiter.array.js",open:"[",close:"]",isElectric:!0},{tokenType:"delimiter.parenthesis.js",open:"(",close:")",isElectric:!0}],docComment:{scope:"comment.doc",open:"/**",lineStart:" * ",close:" */"}}),this.characterPairSupport=new a.CharacterPairSupport(this,{autoClosingPairs:[{open:"{",close:"}"},{open:"[",close:"]"},{open:"(",close:")"},{open:'"',close:'"',notIn:["string"]},{open:"'",close:"'",notIn:["string","comment"]}]}),this.suggestSupport=new a.SuggestSupport(this,{triggerCharacters:["."],excludeTokens:["string","comment","number"],sortBy:[{type:"reference",partSeparator:"/"}],suggest:function(e,t){return n.suggest(e,t)},getSuggestionDetails:function(e,t,r){return n.getSuggestionDetails(e,t,r)}})}return __extends(s,t),s.prototype.asyncCtor=function(){var t=this;return this._threadService.isInMainThread?r.Promise.as(this):new r.Promise(function(r){e(["vs/languages/typescript/typescriptWorker2"],function(){r(t)})})},s.prototype._worker=function(e){return t.prototype._worker.call(this,e)},s.prototype._getProjectResolverDescription=function(){return l.Extensions.getProjectResolver()},s.prototype.getNonWordTokenTypes=function(){return["delimiter.js","delimiter.parenthesis.js","delimiter.bracket.js","delimiter.array.js","regexp.js"]},s.prototype.getCommentsConfiguration=function(){return{lineCommentTokens:["//"],blockCommentStartToken:"/*",blockCommentEndToken:"*/"}},s.prototype.getWordDefinition=function(){return s.JS_WORD_DEFINITION},s.JS_WORD_DEFINITION=n.createWordRegExp("$"),s}(s.TypeScriptMode);t.JSMode=c});
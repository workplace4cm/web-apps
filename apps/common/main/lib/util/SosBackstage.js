/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
/**
 *  SosBackstage.js
 *
 *  Secure Office Suite: the Office-style File menu (Backstage) of
 *  EuroOffice Writer, Sheet, Presentation and PDF in the desktop app.
 *
 *  Each editor's FileMenu calls apply() at the end of applyMode(). The
 *  menu keeps the editor's own commands (Save, Save As, Print, Protect,
 *  Info, Recent, Close, Options) and gets, in Office order:
 *
 *    Home     the application's start page (shell: sos:backstage:home)
 *    New      blank document and the application's templates, the same
 *             ones the start page shows (shell: sos:backstage:templates,
 *             create:new)
 *    Open     the native file picker with the application's formats
 *             (shell: open:folder app:<id>)
 *    Export   formats the engine writes as a separate file, so the open
 *             document stays as it is (the editor's own saveas:format)
 *    Exit     closes the suite; unsaved documents ask first
 *             (shell: sos:backstage:exit)
 *    About    the shell's product name, version and copyright, the same
 *             text as the start page's About (shell: sos:backstage:about)
 *
 *  Commands the desktop app can't run offline (Help, Suggest a Feature)
 *  are hidden, and nothing shows before the shell has answered for it:
 *  About appears once the shell has sent its text.
 */

if (Common === undefined)
    var Common = {};

define([
    'common/main/lib/component/MenuItem'
], function () {
    'use strict';

    Common.UI = Common.UI || {};

    Common.UI.SosBackstage = new (function () {
        var text = function (key, def) {
            var res;
            try {
                res = Common.Locale.get(key, {name: 'Common.Views.SosBackstage', default: def});
            } catch (e) {}
            return res || def;
        };

        var esc = function (s) {
            return Common.Utils.String.htmlEncode(s === undefined || s === null ? '' : String(s));
        };

        var appId = function () {
            return window.PDFE ? 'pdf' : window.SSE ? 'sheet' : window.PE ? 'presentation' : window.DE ? 'writer' : '';
        };

        var native = function () {
            return window.AscDesktopEditor || window.desktop;
        };

        var exec = function (cmd, param) {
            var n = native();
            n && n.execCommand && n.execCommand(cmd, param === undefined ? '{}' : param);
        };

        // the shell opens a window for every create request: ignore double clicks
        var locked = false;
        var once = function (fn) {
            if ( locked ) return;
            locked = true;
            setTimeout(function () { locked = false; }, 1000);
            fn();
        };

        var apps = {
            writer:       {name: 'EuroOffice Writer',       blank: 'word',  blankText: ['txtBlankDocument', 'Blank document'],
                           browse: 'word'},
            sheet:        {name: 'EuroOffice Sheet',        blank: 'cell',  blankText: ['txtBlankWorkbook', 'Blank workbook'],
                           browse: 'cell'},
            presentation: {name: 'EuroOffice Presentation', blank: 'slide', blankText: ['txtBlankPresentation', 'Blank presentation'],
                           browse: 'slide'},
            pdf:          {name: 'EuroOffice PDF',          blank: 'form',  blankText: ['txtBlankForm', 'Blank PDF form']}
        };

        // Export: formats the engine saves as a separate file. Saving an
        // Office document as PDF or images, or a PDF as anything else, writes
        // a new file and leaves the open document on its own file (the
        // desktop engine's LocalFileSaveEnd_CheckSaveAs). Formats that would
        // switch the open document to the new file are Save As, not Export.
        var exportFormats = function (app) {
            var t = Asc.c_oAscFileType,
                office = [
                    {type: t.PDF,  ext: 'pdf', title: ['txtExportPdf', 'PDF document'], desc: ['txtExportPdfDesc', 'Fixed layout for sharing and printing. Fonts and layout stay the same everywhere.']},
                    {type: t.PDFA, ext: 'pdf/a', title: ['txtExportPdfa', 'PDF/A document'], desc: ['txtExportPdfaDesc', 'PDF for long-term archiving.']},
                    {type: t.PNG,  ext: 'png', title: ['txtExportPng', 'PNG images'], desc: ['txtExportImagesDesc', 'A folder with one picture per page.']},
                    {type: t.JPG,  ext: 'jpg', title: ['txtExportJpg', 'JPEG images'], desc: ['txtExportImagesDesc', 'A folder with one picture per page.']}
                ];
            if ( app == 'sheet' ) {
                office[0].desc = ['txtExportPdfSheetDesc', 'Fixed layout for sharing and printing. You choose the sheets and the page setup first.'];
                office[1].desc = ['txtExportPdfaSheetDesc', 'PDF for long-term archiving. You choose the sheets and the page setup first.'];
            }
            if ( app == 'pdf' )
                return [
                    {type: t.DOCX, ext: 'docx', title: ['txtExportDocx', 'Word document (DOCX)'], desc: ['txtExportDocxDesc', 'Editable text recognised from the pages. Complex layouts may change.']},
                    {type: t.ODT,  ext: 'odt', title: ['txtExportOdt', 'OpenDocument text (ODT)'], desc: ['txtExportDocxDesc', 'Editable text recognised from the pages. Complex layouts may change.']},
                    {type: t.RTF,  ext: 'rtf', title: ['txtExportRtf', 'Rich text (RTF)'], desc: ['txtExportRtfDesc', 'Text with basic formatting.']},
                    {type: t.TXT,  ext: 'txt', title: ['txtExportTxt', 'Plain text (TXT)'], desc: ['txtExportTxtDesc', 'Text only, without formatting or pictures.']},
                    office[2], office[3]
                ];
            return office;
        };

        var state = {
            app: '',
            templates: null,        // null until the shell has answered
            userFolder: '',
            about: null,
            menus: []
        };

        // ------------------------------------------------------------------
        // Pages
        // ------------------------------------------------------------------

        var Page = function (id, menu) {
            this.menu = menu;
            this.$el = $('<div id="' + id + '" class="content-box sos-bs-page"></div>').hide();
            menu.$el.find('.panel-context').append(this.$el);
        };
        Page.prototype.show = function () {
            this.render && this.render();
            this.$el.show();
            this.$el.scrollTop(0);
        };
        Page.prototype.hide = function () {
            this.$el.hide();
        };
        Page.prototype.isVisible = function () {
            return this.$el.is(':visible');
        };

        // New: blank document and the application's templates
        var NewPage = function (menu) {
            Page.call(this, 'panel-sos-new', menu);
            var me = this;
            this.$el.on('click', '.sos-bs-card', function () {
                var $card = $(this);
                if ( $card.data('kind') == 'blank' ) {
                    once(function () { exec('create:new', apps[state.app].blank); });
                } else {
                    var tpl = _.findWhere(state.templates || [], {id: $card.data('id')});
                    tpl && once(function () {
                        exec('create:new', JSON.stringify({template: {id: tpl.id, app: tpl.app, type: 0, path: tpl.path}}));
                    });
                }
                me.menu.hide();
            });
            this.$el.on('click', '.sos-bs-browse', function () {
                once(function () { exec('create:new', 'template:' + apps[state.app].browse); });
                me.menu.hide();
            });
            this.$el.on('click', '.sos-bs-folder', function () {
                exec('sos:templates:folder');
            });
        };
        NewPage.prototype = Object.create(Page.prototype);
        NewPage.prototype.show = function () {
            // the list can change while the app runs (My templates)
            exec('sos:backstage:templates', JSON.stringify({app: state.app}));
            Page.prototype.show.call(this);
        };
        NewPage.prototype.render = function () {
            var app = apps[state.app],
                html = '<div class="header">' + esc(text('txtNew', 'New')) + '</div>' +
                       '<div class="sos-bs-grid" role="list">' +
                       '<button type="button" class="sos-bs-card" data-kind="blank" role="listitem">' +
                         '<span class="sos-bs-thumb sos-bs-blank"></span>' +
                         '<span class="sos-bs-name">' + esc(text(app.blankText[0], app.blankText[1])) + '</span>' +
                       '</button>';

            var list = state.templates || [];
            _.each(list, function (t) {
                html += '<button type="button" class="sos-bs-card" data-kind="template" data-id="' + esc(t.id) + '" role="listitem" title="' +
                            esc(t.description || t.name) + '">' +
                          '<span class="sos-bs-thumb">' + (t.thumb ? '<img src="' + esc(t.thumb) + '" alt="" draggable="false">' : '') + '</span>' +
                          '<span class="sos-bs-name">' + esc(t.name) + '</span>' +
                          '<span class="sos-bs-category">' + esc(t.user ? text('txtMyTemplates', 'My templates') : t.category) + '</span>' +
                        '</button>';
            });
            html += '</div>';

            if ( state.templates === null )
                html += '<p class="sos-bs-note">' + esc(text('txtLoadingTemplates', 'Loading templates…')) + '</p>';

            html += '<div class="sos-bs-links">' +
                        (app.browse ? '<button type="button" class="sos-bs-link sos-bs-browse">' + esc(text('txtBrowseTemplate', 'Open a template from this computer…')) + '</button>' : '') +
                        // My templates holds Writer, Sheet and Presentation templates only
                        (app.browse ? '<button type="button" class="sos-bs-link sos-bs-folder">' + esc(text('txtTemplatesFolder', 'Open the My Templates folder')) + '</button>' : '') +
                    '</div>';
            this.$el.html(html);
            // a thumbnail that can't be read leaves a plain card
            this.$el.find('.sos-bs-thumb img').on('error', function () { $(this).remove(); });
        };

        // Export: a separate file in another format
        var ExportPage = function (menu) {
            Page.call(this, 'panel-sos-export', menu);
            var me = this;
            this.$el.on('click', '.sos-bs-format', function () {
                var type = parseInt($(this).data('type'));
                // the editor's own export: per-format warnings and, for Sheet,
                // the page setup; the shell's save dialog offers only this format
                me.menu.fireEvent('saveas:format', [me.menu, type]);
            });
        };
        ExportPage.prototype = Object.create(Page.prototype);
        ExportPage.prototype.render = function () {
            if ( this.rendered ) return;
            this.rendered = true;
            var html = '<div class="header">' + esc(text('txtExport', 'Export')) + '</div>' +
                       '<p class="sos-bs-intro">' + esc(text('txtExportIntro', 'Save a copy in another format. The open document stays as it is.')) + '</p>' +
                       '<div class="sos-bs-formats" role="list">';
            _.each(exportFormats(state.app), function (f) {
                html += '<button type="button" class="sos-bs-format" data-type="' + f.type + '" role="listitem">' +
                          '<span class="sos-bs-ext">' + esc(f.ext.toUpperCase()) + '</span>' +
                          '<span class="sos-bs-format-text">' +
                            '<span class="sos-bs-name">' + esc(text(f.title[0], f.title[1])) + '</span>' +
                            '<span class="sos-bs-desc">' + esc(text(f.desc[0], f.desc[1])) + '</span>' +
                          '</span>' +
                        '</button>';
            });
            html += '</div>';
            this.$el.html(html);
        };

        // About: the shell's text, as on the start page
        var AboutPage = function (menu) {
            Page.call(this, 'panel-sos-about', menu);
        };
        AboutPage.prototype = Object.create(Page.prototype);
        AboutPage.prototype.render = function () {
            var a = state.about || {};
            this.$el.html(
                '<div class="header">' + esc(text('txtAbout', 'About')) + '</div>' +
                '<div class="sos-bs-about">' +
                    '<div class="sos-bs-about-logo"></div>' +
                    '<p class="sos-bs-about-product">' + esc(a.appname) + '</p>' +
                    '<p class="sos-bs-about-app">' + esc(apps[state.app].name) + '</p>' +
                    '<p class="sos-bs-about-version">' + esc(text('txtVersion', 'Version')) + ' ' + esc(a.version) + '</p>' +
                    '<p class="sos-bs-about-rights">' + esc(a.rights) + '</p>' +
                '</div>');
        };

        // ------------------------------------------------------------------
        // Menu items
        // ------------------------------------------------------------------

        var addItem = function (menu, opts, place) {
            var $li = $('<li id="' + opts.id + '" class="fm-btn"></li>');
            place($li);
            var item = new Common.UI.MenuItem({
                el: $li,
                action: opts.action,
                caption: opts.caption,
                canFocused: false,
                dataHint: 1,
                dataHintDirection: 'left-top',
                dataHintOffset: [-2, 22],
                iconCls: opts.icon ? 'menu__icon ' + opts.icon : undefined
            });
            menu.items.push(item);
            return item;
        };

        var find = function (menu, action) {
            return _.find(menu.items, function (item) { return item.options.action == action; });
        };

        var setVisible = function (item, visible) {
            item && item[visible ? 'show' : 'hide']();
        };

        var onItemClick = function (menu, action) {
            switch ( action ) {
            case 'sos:home':
                menu.hide();
                exec('sos:backstage:home', JSON.stringify({app: state.app}));
                break;
            case 'sos:open':
                menu.hide();
                once(function () { exec('open:folder', 'app:' + state.app); });
                break;
            case 'sos:exit':
                menu.hide();
                exec('sos:backstage:exit');
                break;
            }
        };

        var build = function (menu) {
            var sos = menu.sosBackstage = {};
            var $menu = menu.$el.find('.panel-menu');

            sos.home = addItem(menu, {id: 'fm-btn-sos-home', action: 'sos:home', caption: text('txtHome', 'Home'), icon: 'btn-home'},
                function ($li) { $li.insertBefore($menu.find('#fm-btn-create')); });
            sos.newItem = addItem(menu, {id: 'fm-btn-sos-new', action: 'sos:new', caption: text('txtNew', 'New'), icon: 'btn-create-new'},
                function ($li) { $li.insertBefore($menu.find('#fm-btn-create')); });

            var $print = $menu.find('#fm-btn-print-with-preview');
            !$print.length && ($print = $menu.find('#fm-btn-print'));
            sos.exportItem = addItem(menu, {id: 'fm-btn-sos-export', action: 'sos:export', caption: text('txtExport', 'Export'), icon: 'btn-download'},
                function ($li) { $li.insertAfter($print); });

            sos.exit = addItem(menu, {id: 'fm-btn-sos-exit', action: 'sos:exit', caption: text('txtExit', 'Exit')},
                function ($li) {
                    var $close = $menu.find('#fm-btn-exit');
                    $close.length ? $li.insertAfter($close) : $li.insertBefore($menu.find('.devider-last'));
                });

            sos.about = addItem(menu, {id: 'fm-btn-sos-about', action: 'sos:about', caption: text('txtAbout', 'About'), icon: 'btn-menu-about'},
                function ($li) { $li.insertAfter($menu.find('#fm-btn-settings')); });

            menu.panels['sos:new'] = new NewPage(menu);
            menu.panels['sos:export'] = new ExportPage(menu);
            menu.panels['sos:about'] = new AboutPage(menu);

            // a view's own listeners get fireEvent's argument list as one array
            menu.on('item:click', function (args) {
                args = _.isArray(args) ? args : arguments;
                onItemClick(menu, args[1]);
            });
            menu.on('menu:show', function () {
                // a new list of templates for the next visit of New
                exec('sos:backstage:templates', JSON.stringify({app: state.app}));
            });

            // Open: the native picker with this application's formats first
            var open = find(menu, 'file:open');
            open && (open.options.action = 'sos:open');

            exec('sos:backstage:about');
            exec('sos:backstage:templates', JSON.stringify({app: state.app}));
            state.menus.push(menu);
        };

        var refresh = function (menu) {
            var sos = menu.sosBackstage,
                mode = menu.mode || {};

            setVisible(sos.home, true);
            setVisible(sos.newItem, true);
            setVisible(sos.exportItem, !!mode.canDownload);
            setVisible(sos.exit, true);
            setVisible(sos.about, !!state.about);

            // replaced by Home/New/Export above, or not available offline
            setVisible(menu.miNew, false);
            setVisible(find(menu, 'create:fromtemplate'), false);
            setVisible(menu.miExportToPDF, false);
            setVisible(menu.miHelp, false);
            setVisible(menu.miSuggest, false);
            setVisible(menu.miDownload, false);
            setVisible(menu.miSaveCopyAs, false);

            // Office captions
            menu.miRecent && menu.miRecent.setCaption(text('txtRecent', 'Recent'));
            menu.miSaveAs && menu.miSaveAs.setCaption(text('txtSaveAs', 'Save As'));
            menu.miSettings && menu.miSettings.setCaption(text('txtOptions', 'Options'));
            var open = find(menu, 'sos:open');
            open && open.setCaption(text('txtOpen', 'Open'));
            var close = find(menu, 'file:exit');
            close && close.setCaption(text('txtClose', 'Close'));

            // the panels of hidden commands must not stay on screen
            if ( !sos.about.isVisible() && menu.active == 'sos:about' ) menu.active = undefined;
            if ( !sos.exportItem.isVisible() && menu.active == 'sos:export' ) menu.active = undefined;
        };

        var onShellMessage = function (op, param) {
            var data;
            try { data = JSON.parse(param); } catch (e) { return; }

            if ( op == 'templates' ) {
                if ( data.app && data.app != state.app ) return;
                state.templates = _.filter(data.templates || [], function (t) { return t && t.id && t.path && t.name; });
                state.userFolder = data.userFolder || '';
                _.each(state.menus, function (menu) {
                    var page = menu.panels['sos:new'];
                    page && page.isVisible() && page.render();
                });
            } else
            if ( op == 'about' ) {
                if ( !data.appname || !data.version ) return;
                state.about = data;
                _.each(state.menus, function (menu) {
                    refresh(menu);
                    menu.hideDividers && menu.hideDividers();
                    var page = menu.panels['sos:about'];
                    page && page.isVisible() && page.render();
                });
            }
        };

        var listening = false;
        var listen = function () {
            if ( listening ) return;
            listening = true;
            Common.NotificationCenter.on('sos:backstage', onShellMessage);
        };

        return {
            // FileMenu.applyMode(): add the Backstage commands (once) and
            // bring the menu to the Office layout (every time, because
            // applyMode shows the editor's commands again)
            apply: function (menu) {
                if ( !menu || !menu.rendered || !menu.mode || !menu.panels ) return;
                if ( !(menu.mode.isDesktopApp && menu.mode.isOffline) || !native() ) return;

                state.app = state.app || appId();
                if ( !apps[state.app] ) return;

                listen();
                !menu.sosBackstage && build(menu);
                refresh(menu);
            },
            isActive: function () {
                return !!state.app && state.menus.length > 0;
            },
            exportFormats: exportFormats
        };
    })();
});

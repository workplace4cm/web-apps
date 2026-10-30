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
 *  SosRibbon.js
 *
 *  Secure Office Suite: shared pieces of the Office-style ribbon of
 *  EuroOffice Writer, Sheet, Presentation and PDF.
 *
 *  Every application keeps its own tab list (see the editors' Toolbar
 *  controllers). This file holds what they share:
 *  - removing tabs the product does not offer (Plugins),
 *  - folding a tab into another one (Protection and Draw into Review),
 *  - the Help tab,
 *  - building a tab panel out of buttons.
 *
 *  Only commands that run an existing editor operation are added.
 */

if (Common === undefined)
    var Common = {};

define([
    'common/main/lib/component/Button'
], function () {
    'use strict';

    Common.UI = Common.UI || {};

    Common.UI.SosRibbon = new (function () {
        var appPrefix = function () {
            return window.DE ? 'DE' : window.SSE ? 'SSE' : window.PE ? 'PE' : window.PDFE ? 'PDFE' : window.VE ? 'VE' : 'DE';
        };

        var text = function (key, def, scope) {
            var res;
            try {
                res = Common.Locale.get(key, {name: scope || 'Common.Views.SosRibbon', default: def});
            } catch (e) {}
            return res || def;
        };

        // Builds '<section class="panel">' with one group per entry of 'groups';
        // a group is an array of buttons, separators go between groups.
        var createPanel = function (action, groups) {
            var $panel = $('<section class="panel" data-tab="' + action + '" role="tabpanel" aria-labelledby="' + action + '"></section>');
            _.each(groups, function (group, index) {
                if (!group || !group.buttons || !group.buttons.length) return;
                index > 0 && $panel.append('<div class="separator long' + (group.cls ? ' ' + group.cls : '') + '"></div>');
                var $group = $('<div class="group' + (group.cls ? ' ' + group.cls : '') + '"></div>');
                group.label && $group.attr({role: 'group', 'aria-label': group.label});
                _.each(group.buttons, function (btn) {
                    var $slot = $('<span class="btn-slot text x-huge"></span>');
                    $group.append($slot);
                    btn.render($slot);
                });
                $panel.append($group);
            });
            return $panel;
        };

        var createButton = function (opts) {
            var btn = new Common.UI.Button(_.extend({
                cls: 'btn-toolbar x-huge icon-top',
                dataHint: '1',
                dataHintDirection: 'bottom',
                dataHintOffset: 'small'
            }, opts));
            return btn;
        };

        // Drops groups whose buttons were never created (e.g. chat and co-editing when working offline)
        // and the separators that would be left doubled.
        var tidyPanel = function ($panel) {
            if (!$panel || !$panel.length) return;
            $panel.children('.group').each(function () {
                var $slots = $(this).find('.btn-slot');
                if ($slots.length && _.every($slots, function (slot) { return slot.id && !slot.children.length; }))
                    $(this).remove();
            });
            var prevSeparator = true;
            $panel.children().not('.more-box').each(function () {
                if (this.style.display === 'none') return;
                var isSeparator = $(this).hasClass('separator');
                if (isSeparator && prevSeparator) $(this).remove();
                else prevSeparator = isSeparator;
            });
            var $last = $panel.children().not('.more-box').filter(function () { return this.style.display !== 'none'; }).last();
            $last.hasClass('separator') && $last.remove();
        };

        return {
            text: text,
            tidyPanel: tidyPanel,
            createPanel: createPanel,
            createButton: createButton,

            // Tabs that are not part of any Secure Office Suite application.
            removeUnusedTabs: function (toolbar) {
                toolbar.removeTab('plugins');
            },

            // Collaboration -> Review, with Protection and Draw folded in, as in Office.
            buildReviewTab: function (toolbar, opts) {
                opts = opts || {};
                if (!toolbar.getTab('review')) return;
                toolbar.setTabCaption('review', text('textTabReview', 'Review'));
                toolbar.mergeTab('protect', 'review');
                opts.mergeDraw !== false && toolbar.mergeTab('draw', 'review');
                tidyPanel(toolbar.getTab('review'));
            },

            // Help tab: offline help and the keyboard shortcuts list.
            // opts: {mode, getHelp: function () -> bool (opens the help), api}
            addHelpTab: function (toolbar, opts) {
                var me = this,
                    groups = [];
                opts = opts || {};

                var helpButtons = [];
                if (opts.mode && opts.mode.canHelp !== false && opts.showHelp) {
                    var btnHelp = createButton({
                        iconCls: 'toolbar__icon btn-big-help',
                        caption: text('btnHelpCaption', 'Help', appPrefix() + '.Views.FileMenu'),
                        hint: text('tipHelp', 'Open the help')
                    });
                    btnHelp.on('click', function () {
                        opts.showHelp();
                    });
                    helpButtons.push(btnHelp);
                }

                var btnShortcuts = createButton({
                    iconCls: 'toolbar__icon btn-big-keyboard',
                    caption: text('txtTitle', 'Keyboard Shortcuts', 'Common.Views.ShortcutsDialog'),
                    hint: text('tipShortcuts', 'Show and change keyboard shortcuts')
                });
                btnShortcuts.on('click', function () {
                    if (!Common.Views.ShortcutsDialog) return;
                    if (me.dlgShortcuts && me.dlgShortcuts.isVisible()) return;
                    me.dlgShortcuts = new Common.Views.ShortcutsDialog({api: opts.api});
                    me.dlgShortcuts.show();
                });
                helpButtons.push(btnShortcuts);

                groups.push({buttons: helpButtons, label: text('textGroupHelp', 'Help')});

                var $panel = createPanel('sos-help', groups);
                toolbar.addTab({action: 'sos-help', caption: text('textTabHelp', 'Help'), dataHintTitle: 'Y', layoutname: 'toolbar-sos-help'},
                                $panel, toolbar.getLastTabIdx());
                toolbar.setVisible('sos-help', true);
                return $panel;
            }
        };
    })();
});

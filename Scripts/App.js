/// <reference path="_references.js" />
'use strict';

$(function () {
    var app = {
        getUrlParamByName: function (name) {
            name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
            var regex = new RegExp("[\\?&]" + name + "=([^&#]*)");
            var results = regex.exec(location.search);
            return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
        },
        appWebUrl: function () {
            var value = this.getUrlParamByName("SPAppWebUrl");
            return {
                Full: value,
                Relative: "/" + value.replace(/^(?:\/\/|[^\/]+)*\//, "")
            }
        },

        hostWebUrl: function () {
            var value = this.getUrlParamByName("SPHostUrl");
            return {
                Full: value,
                Relative: "/" + value.replace(/^(?:\/\/|[^\/]+)*\//, "")
            }
        },
        applyBrandingCA: function () {
            // get the current host and app web urls.
            var hostWebUrl = this.hostWebUrl();
            var appWebUrl = this.appWebUrl();

            // Upload the file
            $pnp.sp.web.getFileByServerRelativeUrl(appWebUrl.Relative + "/Content/branding.css").getBuffer().then(function (buffer) {
                $pnp.sp.crossDomainWeb(appWebUrl.Full, hostWebUrl.Full).getFolderByServerRelativeUrl(hostWebUrl.Relative + "/SiteAssets").files.add("branding.css", buffer, true).then(function (result) {

                    // Create the CSS link
                    var str = [
                        "var linkDemo = document.createElement('LINK');",
                        "linkDemo.type = 'text/css';",
                        "linkDemo.rel = 'stylesheet';",
                        "linkDemo.href = '" + hostWebUrl.Full + "/SiteAssets/branding.css';",
                        "document.getElementsByTagName('head')[0].appendChild(linkDemo);"
                    ].join("");

                    // Apply Custom Action
                    $pnp.sp.site.userCustomActions.add({
                        Title: "SharePoint Fest Branding Demo",
                        Name: "SharePoint Fest Branding Demo",
                        Description: "Apply custom branding for SharePoint Fest Demo!",
                        Location: "ScriptLink",
                        ScriptBlock: str,
                        Sequence: 10011
                    }).then(function () {
                        $("#msg").text("Branding Applied using a Custom Action!");
                    });
                });
            });
        },
        removeBrandingCA: function () {
            // get the current host and app web urls.
            var hostWebUrl = this.hostWebUrl();
            var appWebUrl = this.appWebUrl();

            // Remove Custom Action
            $pnp.sp.site.userCustomActions.get().then(function (res) {
                for (var i = 0; i < res.length; i++) {
                    if (res[i].Name == "SharePoint Fest Branding Demo") {
                        $pnp.sp.site.userCustomActions.getById(res[i].Id).delete();
                        break;
                    }
                }
                // Delete the file
                $pnp.sp.crossDomainWeb(appWebUrl.Full, hostWebUrl.Full).getFileByServerRelativeUrl(hostWebUrl.Relative + "/SiteAssets/branding.css").delete().then(function () {
                    $("#msg").text("Branding has been removed... :(");
                });
            });
        },
        applyBrandingAltCss: function () {
            // get the current host and app web urls.
            var hostWebUrl = this.hostWebUrl();
            var appWebUrl = this.appWebUrl();

            // Upload the file
            $pnp.sp.web.getFileByServerRelativeUrl(appWebUrl.Relative + "/Content/branding.css").getBuffer().then(function (buffer) {
                $pnp.sp.crossDomainWeb(appWebUrl.Full, hostWebUrl.Full).getFolderByServerRelativeUrl(hostWebUrl.Relative + "/SiteAssets").files.add("branding.css", buffer, true).then(function (result) {
                    // Apply AlternateCSS
                    $pnp.sp.crossDomainWeb(appWebUrl.Full, hostWebUrl.Full).update({ AlternateCssUrl: hostWebUrl.Full + "/SiteAssets/branding.css" }).then(function () {
                        $("#msg").text("Branding Applied using AlternateCSS!");
                    });
                });
            });
        },
        removeBrandingAltCss: function () {
            // get the current host and app web urls.
            var hostWebUrl = this.hostWebUrl();
            var appWebUrl = this.appWebUrl();

            // Remove AlternateCSS
            $pnp.sp.crossDomainWeb(appWebUrl.Full, hostWebUrl.Full).update({ AlternateCssUrl: "" }).then(function () {
                // Delete the file
                $pnp.sp.crossDomainWeb(appWebUrl.Full, hostWebUrl.Full).getFileByServerRelativeUrl(hostWebUrl.Relative + "/SiteAssets/branding.css").delete().then(function () {
                    $("#msg").text("Branding has been removed... :(");
                });
            });
        }
    };

    // Button events
    $("#btn1").on("click", app.applyBrandingCA.bind(app));
    $("#btn2").on("click", app.removeBrandingCA.bind(app));
    $("#btn3").on("click", app.removeBrandingAltCss.bind(app));
    $("#btn4").on("click", app.removeBrandingAltCss.bind(app));
});
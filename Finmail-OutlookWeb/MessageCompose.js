(function () {
    "use strict";

    const appdomain = 'https://app.finmail.com';

    const zh = {
        'OK': "成功",
        'Failed': "失败",
        'Confirming': "确认中",
        'Address copied': "地址已复制",
        'Failed to copy address': "地址复制失败",
        'Invalid address': "无效地址",
        'Failed to create or change Finmail password. System error.': "无法创建或更改风邮密码。系统错误。",
        'Failed to create or change backup email.System error.': "无法创建或更改备份邮箱。系统错误。",
        'Failed to update Finmail password. Backup email should not be the one in Outlook.': "备份邮箱不应与Outlook邮箱相同",
        'The input is invalid. Please try again.': "输入错误，请重试。",
        'Failed to change Finmail password. System error.': "无法更改风邮密码。系统错误。",
        'Finmail password has been updated. Please check backup email.': "风邮密码更新成功，请检查备份邮箱。",
        'Finmail password has been reset. Please check backup email.': "风邮密码重置成功，请检查备份邮箱。",
        'Failed to reset Finmail password': "无法重置风邮密码",
        'Failed to reset Finmail password. System error.': "无法重置风邮密码。系统错误。",
        'Too many update or reset operations. Please try tomorrow.': "更新或重置操作过多。请明天再试。",
        'Backup email has not been verified': "备份邮箱尚未验证",
        'Failed to send payment': "汇款失败",
        'Failed to get user information': "无法得到用户信息",
        'Failed to get user information. System error.': "无法得到用户信息。系统错误。",
        'Failed to get or create user information.': "无法得到或创建用户信息。系统错误。",
        'Failed to get new user deposit address': "无法得到新用户存款地址",
        'Failed to create new user': "无法创建新用户",
        'Failed to get user transaction history': "无法得到用户存/汇款历史",
        'Unsupported method': "不支持此方式",
        'Invalid parameter': "参数无效",
        'Invalid user': "无效用户",
        'Invalid backup email': "备份邮箱无效",
        'Invalid user email': "当前邮箱无效",
        'Invalid current Finmail password': "风邮密码错误",
        'Failed to update Finmail password': "无法更改风邮密码",
        'Failed to update Finmail password. System error.': "无法更改风邮密码。系统错误。",
        'Payout amount exceeds daily limit': "汇款额超过每日限额",
        'Invalid Finmail password': "风邮密码错误",
        'Insufficient balance': "余额不足",
        'Failed to create transaction': "汇款失败",
        'Failed to create transaction. System error.': "汇款失败。系统错误。",
        'Invalid transaction': "交易无效",
        'Failed to execute transaction': "无法进行汇款",
        'Failed to execute transaction. System error.': "无法进行汇款 。系统错误。",
        'Unable to validate user': "无法验证用户",
        'Invalid BTC address. Please try again.': "BTC地址无效，请重试",
        'Invalid amount. Please try again.': "金额无效，请重试",
        'Invalid fee. Please try again.': "费用无效，请重试",
        'Insufficient balance. Please try again.': "余额不足，请重试",
        'Invalid Finmail password. Please try again.': "风邮密码无效，请重试",
        'Date': "日期",
        'Amount': "金额",
        'Status': "状态",
        'Address': "地址",
        'Fee': "费用",
        'Create new Finmail password': "创建新风邮密码",
        'Change Finmail password': "更改风邮密码",
        'Create': "创建",
        'Change': "更改"
    }

    const email_html_template1 = '<p><font color="#ffa500">The email sender has made a new payment of</font> <font color="blue">(#amount)' +
        ' (#currency_name)</font> <font color="#ffa500"> to</font> <font color="blue">' +
        '(#address)</font><font color="#ffa500">(#appex_2).' +
        ' The <font color = "blue">(#transaction_type)' +
        '</font><font color="#ffa500"> transaction ID is</font > <font color="blue">(#transaction_id)' +
        '</font><font color="#ffa500">. Please use <a href="https://www.finmail.com" target="_blank">Finmail Add-in</a>(#appex_1) to review the payment.</p>'
    const email_text_template1 = 'The email sender has made a new payment of (#amount)' +
        ' (#currency_name) to (#address)(#appex_2). The (#transaction_type) transaction ID is ' +
        ' (#transaction_id). Please use Finmail Add-in at https://www.finmail.com(#appex_1) to review the payment.';

    const pwd_length = 16;
    var currencies = null;
    var deposit_history = null;
    var payout_history = null;

    var $panel;
    var $panelMain;
    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('.ms-Pivot').initPivot();
            $('.ms-Panel').initPanel();
            $('#currency_dropdown').initDropdown();
            $('#fee_dropdown').initDropdown();

            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            //messageBanner._hideBanner();
            loadProps();
            getUserInfo(true);
            setInterval(getUserInfo, 5000);

            $(".start_profile").on('click', function () {
                $('#profile_panel_pivot').trigger("click");
            });

            $("#amount").keyup(function () {
                var amount = parseFloat($('#amount').val());
                var currency = getCurrency();
                var str_amount_value;

                if (currency) {
                    if (amount) {
                        str_amount_value = (amount * currency.value).toFixed(2);
                        $('#amount_label').text("Amount (" + currency.name + ", ~" + str_amount_value + " USD)");
                        $('#panel_amount_label').text("Amount (" + currency.name + ", ~" + str_amount_value + " USD):");
                    } else {
                        $('#amount_label').text("Amount (" + currency.name + ")");
                        $('#panel_amount_label').text("Amount (" + currency.name + "):");
                    }
                }
            });

            $("#currency").change(function () {
                $('#fee_dropdown').initDropdown("refresh");
                updateFinInfoUI();
            });

            $("#fee").change(function () {
                var currency = getCurrency();

                if (currency) {
                    var str_fee_1h_value = (currency.fee_1h * currency.value).toFixed(2);
                    var str_fee_2h_value = (currency.fee_2h * currency.value).toFixed(2);
                    var str_fee_6h_value = (currency.fee_6h * currency.value).toFixed(2);
                    var str_fee_12h_value = (currency.fee_12h * currency.value).toFixed(2);
                    var fee_1h_text = "< 1 Hour, " + currency.fee_1h + " " + currency.name + ", ~" + str_fee_1h_value + " USD";
                    var fee_2h_text = "< 2 Hours, " + currency.fee_2h + " " + currency.name + ", ~" + str_fee_2h_value + " USD";
                    var fee_6h_text = "< 6 Hours, " + currency.fee_6h + " " + currency.name + ", ~" + str_fee_6h_value + " USD";
                    var fee_12h_text = "< 12 Hours, " + currency.fee_12h + " " + currency.name + ", ~" + str_fee_12h_value + " USD";

                    $('#fee_1h').text(fee_1h_text);
                    $('#fee_2h').text(fee_2h_text);
                    $('#fee_6h').text(fee_6h_text);
                    $('#fee_12h').text(fee_12h_text);
                }

                $('#fee_dropdown').initDropdown("refresh");

                var selected_text = $('#fee option:selected').text();
                $('#panel_fee').text(selected_text);

                $('#fee_label').text("Speed & Fee (" + currency.name + ", no fee for internal transfer if applicable)");
                $('#panel_fee_label').text("Speed & Fee (" + currency.name + ", no fee for internal transfer if applicable):");
            });

            $("#get_started").on('click', function () {
                $('#pay_panel_pivot').trigger("click");
            });

            $("#pay").on('click', function () {
                var address = $('#address').val();
                var amount = parseFloat($('#amount').val());
                var currency = getCurrency();
                var fee = getFee(currency);

                if (currency) {
                    // Do not need to check fee here, it will be handled by the server
                    if (!(address && address.length > 0)) {
                        showNotification(lang("Invalid " + currency.name + " or Email address. Please try again."));
                    } else if (!(amount > 0)) {
                        showNotification(lang("Invalid amount. Please try again."));
                    } else if (!(amount + fee <= currency.balance)) {
                        showNotification(lang("Insufficient balance. Please try again."));
                    } else {
                        $("#panel_address").text(address);
                        $("#panel_amount").text(amount);

                        $('#amount').trigger("keyup");
                        $('#fee').trigger("change");

                        // Panel must be set to display "block" in order for animations to render
                        $("#panel_main_pay").css({ display: "block" });
                        $("#panel_pay").addClass("is-open");
                    }
                } else {
                    showNotification(lang("Insufficient currency."));
                }
            });

            $("#confirm").on('click', function () {
                var pwd;

                $("#panel_main_pay").css({ display: "none" });
                $("#panel_pay").removeClass("is-open");

                pwd = $('#pwd').val();
                $('#pwd').val("");
                $('#pwd').focus();

                performPayment(pwd);
            });

            $("#cancel").on('click', function () {
                $("#panel_main_pay").css({ display: "none" });
                $("#panel_pay").removeClass("is-open");

                $('#pwd').val("");
                $('#pwd').focus();
            });

            $("#panel_transaction_update_email").on('click', function () {
                var item = Office.context.mailbox.item;

                if (item.body.getTypeAsync) {
                    item.body.getTypeAsync(function (result) {
                        if (result.status == Office.AsyncResultStatus.Failed) {
                            showNotification(lang("Failed to update email"));
                        } else {
                            var amount = $('#panel_transaction_amount').text();
                            var currency_name = $('#panel_transaction_currency').text();
                            var address = $('#panel_transaction_receiver_address').text();
                            var transaction_type = $('#panel_transaction_type').text();
                            var transaction_id = $('#panel_transaction_id').text();
                            var transaction_status = $('#panel_transaction_status').text();
                            var str;

                            if (result.value == Office.MailboxEnums.BodyType.Html) {
                                //var str = '<p><font color="#ffa500">发件人新汇出一笔</font> <font color="blue">' + amount +
                                //    ' BTC</font> <font color="#ffa500">的款项至</font> <font color="blue">' +
                                //    address + '</font><font color="#ffa500">。交易号为</font> <font color="blue">' +
                                //    txid + '</font><font color="#ffa500">。</font></p><br />';
                                str = email_html_template1;
                            } else {
                                str = email_text_template1;
                            }

                            str = str.replace('(#amount)', amount)
                                .replace('(#currency_name)', currency_name)
                                .replace('(#address)', address)
                                .replace('(#transaction_type)', transaction_type)
                                .replace('(#transaction_id)', transaction_id);
                            if (transaction_type != "External") {
                                str = str.replace('(#appex_1)', '')
                                    .replace('(#appex_2)', '');
                            } else {
                                str = str.replace('(#appex_1)', ' or block explorer');
                                if (transaction_status == "Cancelled") {
                                    str = str.replace('(#appex_2)', ', but has been <font color="blue">Cancelled</font>');
                                } else {
                                    str = str.replace('(#appex_2)', '');
                                }
                            }
                            if (result.value == Office.MailboxEnums.BodyType.Html) {
                                str += '</font></p><br />';
                                item.body.prependAsync(str,
                                    { coercionType: Office.CoercionType.Html });
                            } else {
                                str += '\r\n';
                                item.body.prependAsync(str,
                                    { coercionType: Office.CoercionType.Text });
                            }
                        }
                    })
                }
            });

            $("#change_pwd").on('click', function () {
                var current_pwd;
                var new_pwd;
                var confirm_new_pwd;
                var new_backup_email;

                current_pwd = $('#current_pwd').val();
                $('#current_pwd').val("");

                new_pwd = $('#new_pwd').val();
                $('#new_pwd').val("");
                confirm_new_pwd = $('#confirm_new_pwd').val();
                $('#confirm_new_pwd').val("");
                new_backup_email = $('#backup_email').val();
                $('#backup_email').val("");

                $('#current_pwd').focus();

                performChangePassword(current_pwd, new_pwd, confirm_new_pwd, new_backup_email);
            });

            $("#reset_pwd").on('click', function () {
                performResetPassword(false);
            });

            $("#forgot_pwd").on('click', function () {
                performResetPassword(true);
            });

            $("#payment_reset_pwd").on('click', function () {
                $("#panel_main_pay").css({ display: "none" });
                $("#panel_pay").removeClass("is-open");

                $('#pwd').val("");
                $('#pwd').focus();

                performResetPassword(false);
            });

            $("#sign_up").on('click', function () {
                var sign_up_pwd;
                var confirm_sign_up_pwd;
                var sign_up_backup_email;

                sign_up_pwd = $('#sign_up_pwd').val();
                $('#sign_up_pwd').val("");
                confirm_sign_up_pwd = $('#confirm_sign_up_pwd').val();
                $('#confirm_sign_up_pwd').val("");
                sign_up_backup_email = $('#sign_up_backup_email').val();
                $('#sign_up_backup_email').val("");

                $('#sign_up_pwd').focus();

                performUserSignUp(sign_up_pwd, confirm_sign_up_pwd, sign_up_backup_email);
            });

            $("#sign_in").on('click', function () {
                var sign_in_pwd;

                sign_in_pwd = $("#sign_in_pwd").val();
                $("#sign_in_pwd").val("");
                $("#sign_in_pwd").focus();

                if (!sign_in_pwd) {
                    $("#sign_in_msg").text("Password can not be empty");
                } else {
                    performUserSignIn(sign_in_pwd);
                }
            });

            $("#sign_out").on('click', function () {
                performUserSignOut();
            });
        })
    };

    function lang(str) {
        var lang_ary = {};

        switch (Office.context.displayLanguage) {
            case 'zh-CN':
                //lang_ary = zh;
                break;
            default:
            // To be updated
            //lang_ary = zh;
        }
        if (lang_ary[str]) {
            return lang_ary[str];
        } else {
            return str;
        }
    }

    // Load the message-specific properties.
    function loadProps() {
    }

    function updateProps(address, amount, txid) {
        var item = Office.context.mailbox.item;
        item.internetHeaders.getAsync(["Pay-To"],
            function (asyncResult) {
                var msg;
                // TODO, limit msg length(<1000)
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    if (asyncResult.value) {
                        msg = asyncResult.value["Pay-To"];
                    }
                    if (!msg) {
                        msg = "";
                    }
                    if (msg.length > 0) {
                        msg += ", ";
                    }
                    msg += "<" + address + ">; amount=" + amount +
                        "; unit=btc; txid=<" + txid + ">";
                    item.internetHeaders.setAsync({ "Pay-To": msg });
                }
            }
        );
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        //messageBanner.toggleExpansion();
    }

    function performChangePassword(current_pwd, new_pwd, confirm_new_pwd, new_backup_email, cb) {
        var current_md;
        var new_md;
        var hashedCurrentPwd;
        var hashedNewPwd;
        var new_pwd_is_valid;
        var confirm_new_pwd_is_valid;
        var new_backup_email_is_valid;
        var new_pwd_only;
        var email_only;
        var both_new_pwd_and_email;
        var change_pwd;

        new_pwd_is_valid = (new_pwd && new_pwd.length <= pwd_length);
        confirm_new_pwd_is_valid = (confirm_new_pwd && confirm_new_pwd.length <= pwd_length);
        new_backup_email_is_valid = validateEmail(new_backup_email);

        new_pwd_only = (new_pwd_is_valid && confirm_new_pwd_is_valid && new_pwd == confirm_new_pwd &&
            !new_backup_email);
        email_only = (!new_pwd && !confirm_new_pwd && new_backup_email_is_valid);
        both_new_pwd_and_email = (new_pwd_is_valid && confirm_new_pwd_is_valid && new_pwd == confirm_new_pwd &&
            new_backup_email_is_valid);

        if (new_pwd_only || email_only || both_new_pwd_and_email) {
            change_pwd = (current_pwd && current_pwd.length <= pwd_length)

            if (change_pwd) {
                current_md = new KJUR.crypto.MessageDigest({
                    alg: "sha256",
                    prov: "cryptojs"
                });

                new_md = new KJUR.crypto.MessageDigest({
                    alg: "sha256",
                    prov: "cryptojs"
                });

                if (current_md && new_md) {
                    hashedNewPwd = new_md.digestString(new_pwd);

                    hashedCurrentPwd = current_md.digestString(current_pwd);
                    if (new_pwd_only) {
                        new_backup_email = null;
                    } else if (email_only) {
                        new_pwd = null;
                        hashedNewPwd = null;
                    }

                    Office.context.mailbox.getUserIdentityTokenAsync(performChangePasswordCb,
                        {
                            'new-pwd': new_pwd,
                            'hashed-current-pwd': hashedCurrentPwd,
                            'hashed-new-pwd': hashedNewPwd,
                            'backup-email': new_backup_email,
                            'cb': cb
                        });
                } else {
                    showNotification(lang("Failed to create or change Finmail password. System error."));
                }
            } else {
                showNotification(lang("The input is invalid. Please try again."));
            }
        } else {
            showNotification(lang("The input is invalid. Please try again."));
        }
    }

    function performChangePasswordCb(asyncResult) {
        var endpoint = appdomain + '/auth';
        var token = asyncResult.value;
        var data = asyncResult.asyncContext;
        var profile = Office.context.mailbox.userProfile;
        var newPwd = data["new-pwd"];
        var hashedCurrentPwd = data["hashed-current-pwd"];
        var hashedNewPwd = data["hashed-new-pwd"];
        var backupEmail = data["backup-email"];
        var cb = data.cb;

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'method': "change",
                'token': token,
                'type': "msexch",
                'hashed-current-pwd': hashedCurrentPwd,
                'hashed-new-pwd': hashedNewPwd,
                'backup-email': backupEmail
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                showNotification(lang(data.result));
                if (cb) {
                    cb(newPwd);
                }
            },
            error: function (xhr, status, error) {
                // error
                showNotification(lang("Failed"));
            }
        }).done(function (data) {
        });
    }

    function performResetPassword(sign_in_flag) {
        Office.context.mailbox.getUserIdentityTokenAsync(performResetPasswordCb,
            { 'sign-in-flag': sign_in_flag });
    }

    function performResetPasswordCb(asyncResult) {
        var endpoint = appdomain + '/auth';
        var token = asyncResult.value;
        var profile = Office.context.mailbox.userProfile;
        var sign_in_flag = asyncResult.asyncContext["sign-in-flag"];
        var msg;

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'method': "reset",
                'token': token,
                'type': "msexch"
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                if (sign_in_flag) {
                    $("#sign_in_msg").text(lang(data.result));
                } else {
                    showNotification(lang(data.result));
                }
            },
            error: function (xhr, status, error) {
                // error
                msg = lang("Failed to reset Finmail password due to network issue");

                if (sign_in_flag) {
                    $("#sign_in_msg").text(msg);
                } else {
                    showNotification(msg);
                }
            }
        }).done(function (data) {
        });
    }

    function performPayment(pwd) {
        var currency_id;
        var currency;
        var address;
        var amount;
        var fee;
        var fee_btc;
        var md;
        var hashedPwd;

        currency_id = $('#currency option:selected').index();
        currency = getCurrency();
        address = $('#address').val();
        amount = parseFloat($('#amount').val());
        fee = getFee(currency);

        $('#address').focus();

        if (currency) {
            if (currency_id == 0) {
                // BTC
                fee_btc = fee;
            } else if (currency_id == 1) {
                // USDT
                fee_btc = fee * currency.value / currencies[0].value;
            } else {
                // Should not happen
                fee_btc = -1;
            }

            if (!(address && address.length > 0)) {
                showNotification(lang("Invalid " + currency.name + " or Email address. Please try again."));
            } else if (!(amount > 0)) {
                showNotification(lang("Invalid amount. Please try again."));
            } else if (!(fee_btc <= 0.003)) {
                showNotification(lang("Invalid fee. Please try again."));
            } else if (!(amount + fee <= currency.balance)) {
                showNotification(lang("Insufficient balance. Please try again."));
            } else if (!(pwd && pwd.length <= 16)) {
                showNotification(lang("Invalid Finmail password. Please try again."));
            } else {
                md = new KJUR.crypto.MessageDigest({
                    alg: "sha256",
                    prov: "cryptojs"
                });

                if (md) {
                    hashedPwd = md.digestString(pwd);
                    Office.context.mailbox.getUserIdentityTokenAsync(performPaymentCb,
                        {
                            'currency-name': currency.name,
                            'address': address,
                            'amount': amount,
                            'fee': fee,
                            'hashed-pwd': hashedPwd
                        });
                } else {
                    showNotification(lang("Failed to change Finmail password. System error."));
                }
            }
        } else {
            showNotification(lang("Invalid currency"));
        }
    }

    function performPaymentCb(asyncResult) {
        var item = Office.context.mailbox.item;
        var endpoint = appdomain + '/payments';
        var token = asyncResult.value;
        var data = asyncResult.asyncContext;
        var currency_name = data["currency-name"];
        var address = data.address;
        var amount = data.amount;
        var fee = data.fee;
        var hashedPwd = data["hashed-pwd"];
        var profile = Office.context.mailbox.userProfile;
        var currency = getCurrency();

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'method': "prepare",
                'token': token,
                'type': "msexch",
                'hashed-pwd': hashedPwd,
                'currency-name': currency_name,
                'address': address,
                'amount': amount,
                'fee': fee,
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                if (data.result == "OK") {
                    var txid = data.txid;

                    //TODO, check api avalibality and,
                    //sync with the follow-up callings
                    //updateProps(address, amount, txid);

                    $.ajax({
                        url: endpoint,
                        crossDomain: true,
                        data: JSON.stringify({
                            'email': profile.emailAddress,
                            'method': "send",
                            'token': token,
                            'type': "msexch",
                            'currency-name': currency_name,
                            'txid': txid
                        }),
                        // headers: {'X-Requested-With': 'XMLHttpRequest'},
                        contentType: 'application/json; charset=utf-8',
                        type: 'POST',
                        dataType: 'json',
                        success: function (data) {
                            var str;
                            if (data.result == "OK") {
                                currency.balance = data.balance;
                                currency.frozen_payout_balance = data.frozen_payout_balance;
                                currency.daily_paid_amount = data.daily_paid_amount;
                                updateFinInfoUI();

                                if (item.body.getTypeAsync) {
                                    item.body.getTypeAsync(function (result) {
                                        if (result.status == Office.AsyncResultStatus.Failed) {
                                            showNotification(lang("Failed to update message"));
                                        } else {
                                            showNotification(lang(data.result));

                                            if (result.value == Office.MailboxEnums.BodyType.Html) {
                                                //var str = '<p><font color="#ffa500">发件人新汇出一笔</font> <font color="blue">' + amount +
                                                //    ' BTC</font> <font color="#ffa500">的款项至</font> <font color="blue">' +
                                                //    address + '</font><font color="#ffa500">。交易号为</font> <font color="blue">' +
                                                //    txid + '</font><font color="#ffa500">。</font></p><br />';
                                                str = email_html_template1;
                                            } else {
                                                str = email_text_template1;
                                            }

                                            str = str.replace('(#amount)', amount)
                                                .replace('(#currency_name)', currency_name)
                                                .replace('(#address)', address)
                                                .replace('(#transaction_type)', data.type)
                                                .replace('(#transaction_id)', txid);

                                            if (data.type != "External") {
                                                str = str.replace('(#appex_1)', '')
                                                    .replace('(#appex_2)', '');
                                            } else {
                                                // The payout transaction including USDT should always be valid
                                                str = str.replace('(#appex_1)', ' or block explorer')
                                                    .replace('(#appex_2)', '');
                                            }

                                            if (result.value == Office.MailboxEnums.BodyType.Html) {
                                                str += '</font></p><br />';
                                                item.body.prependAsync(str,
                                                    { coercionType: Office.CoercionType.Html });
                                            } else {
                                                str += '\r\n';
                                                item.body.prependAsync(str,
                                                    { coercionType: Office.CoercionType.Text });
                                            }
                                        }
                                    })
                                } else {
                                    showNotification(lang(data.result));
                                }
                            }
                        },
                        error: function (xhr, status, error) {
                            // error
                            showNotification(lang("Failed to send payment"));
                        }
                    }).done(function (data) {
                    });
                } else {
                    showNotification(lang(data.result));
                }
            },
            error: function (xhr, status, error) {
                // error
                showNotification(lang("Failed to send payment"));
            }
        }).done(function (data) {
        });
    }

    function updateFinInfoUI() {
        var currency;
        var str_value;
        var infoHtml;

        currency = getCurrency();
        if (currency) {
            str_value = currency.value.toFixed(2);

            $('#pay_panel_messagebar').removeClass('ms-MessageBar--warning');
            $('#pay_panel_messagebar').addClass('ms-MessageBar--success');
            $('#pay_panel_messagebar_icon').removeClass('ms-Icon--refresh');
            $('#pay_panel_messagebar_icon').addClass('ms-Icon--receiptCheck');

            infoHtml = "Balance: " + currency.balance + " " + currency.name + "<br /> ";
            infoHtml += "Confirming (Payout / Deposit): " +
                currency.frozen_payout_balance + " / " + currency.frozen_deposit_balance +
                " " + currency.name + "<br /> ";
            infoHtml += "Today's Payout Limit: " + currency.daily_paid_amount + " / " +
                currency.daily_payout_limit + " " + currency.name + "<br />";
            infoHtml += "1 " + currency.name + " = " + str_value + " USD<br /><br />"
            infoHtml += "Deposit Address (" + currency.name + "):<br />" +
                '<span style="cursor: pointer; " id="deposit_address">' + currency.address +
                "</span>";

            $('#info').html(infoHtml);

            $('#address_label').text("Receiver Address (" + currency.name + " or Email)");
            $('#panel_address_label').text("Receiver Address (" + currency.name + " or Email):");

            $('#amount').trigger("keyup");
            $('#fee').trigger("change");

            $("#deposit_address").on('click', function () {
                var copy = function (e) {
                    var deposit_address = $('#deposit_address').text();
                    var copied;

                    e.preventDefault();
                    if (deposit_address && deposit_address.length > 0) {
                        if (e.clipboardData) {
                            e.clipboardData.setData("text/plain", deposit_address);
                            copied = true;
                        } else if (window.clipboardData) {
                            window.clipboardData.clearData();
                            copied = window.clipboardData.setData("Text", deposit_address);
                        }
                    }

                    if (copied == true) {
                        showNotification(lang("Address copied"));
                    } else {
                        showNotification(lang("Failed to copy address"));
                    }
                }

                window.addEventListener('copy', copy);
                document.execCommand('copy');
                window.removeEventListener('copy', copy);
            });
        }
    }

    function showMainPanel(show_flag) {
        if (show_flag) {
            $("#main_panel").css({ display: "block" });
        } else {
            $("#main_panel").css({ display: "none" });
        }
    }

    function showSignInPanel(show_flag) {
        if (show_flag) {
            // Panel must be set to display "block" in order for animations to render
            $("#panel_main_sign_in").css({ display: "block" });
            $("#panel_sign_in").addClass("is-open");
        } else {
            $("#panel_main_sign_in").css({ display: "none" });
            $("#panel_sign_in").removeClass("is-open");
        }
    }

    function showSignUpPanel(show_flag) {
        if (show_flag) {
            // Panel must be set to display "block" in order for animations to render
            $("#panel_main_sign_up").css({ display: "block" });
            $("#panel_sign_up").addClass("is-open");
        } else {
            $("#panel_main_sign_up").css({ display: "none" });
            $("#panel_sign_up").removeClass("is-open");
        }
    }

    function showPanel(main_panel, sign_up_panel, sign_in_panel) {
        showMainPanel(main_panel);
        showSignUpPanel(sign_up_panel);
        showSignInPanel(sign_in_panel);
    }

    function getUserInfo(show_sign_in) {
        Office.context.mailbox.getUserIdentityTokenAsync(getUserInfoCb,
            {
                'show_sign_in': show_sign_in
            });
    }

    function getUserInfoCb(asyncResult) {
        var token = asyncResult.value;
        var endpoint = appdomain + '/users';
        var profile = Office.context.mailbox.userProfile;
        var date;
        var status;
        var method = 'query';
        var show_sign_in = asyncResult.asyncContext['show_sign_in'];

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'token': token,
                'type': "msexch",
                'method': method,
                'name': profile.displayName
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                if (data.result == "OK") {
                    if (data.user_status == "Waiting for activation" ||
                        data.user_status == "Locked") { // To be updated in future
                        showPanel(false, true, false);
                    } else if (data.user_status == "Signed out") {
                        showPanel(false, false, true);
                    } else if (show_sign_in) {
                        showPanel(false, false, true);
                    } else {
                        currencies = data.currencies;
                        updateFinInfoUI();

                        deposit_history = data.deposit_transactions;
                        if (deposit_history.length > 0) {
                            var trHTML = '<table class="ms-Table"><thead><tr><th>' + lang("Date") +
                                '</th > <th>' + lang("Currency") + '</th > <th>' + lang("Amount") +
                                '</th> <th>' + lang("Status") + '</th></tr ></thead > <tbody>';
                            $.each(deposit_history, function (i, item) {
                                date = new Date(item.time).toLocaleDateString();
                                status = item.status;
                                if (status == "Confirmed") {
                                    status = lang("OK");
                                } else {
                                    status = lang("Confirming") + " (" + item.confirmations + '/3)';
                                }
                                trHTML += '<tr><td>' + date + '</td><td>' + item.currency_name +
                                    '</td><td>' + item.amount + '</td><td>' +
                                    status + '</td></tr>';
                            });
                            trHTML += '</tbody></table>';
                            $('#deposit_history').html(trHTML);

                            $('#deposit_history td').click(function () {
                                var id = $(this).parent().index();
                                var item = deposit_history[id];
                                var time = new Date(item.time).toLocaleString();

                                $('#panel_transaction_receiver_address_label').text("Receiver Address (" + item.currency_name + " or Email):");
                                $('#panel_transaction_amount_label').text("Amount (" + item.currency_name + "):");
                                $('#panel_transaction_fee_label').text("Fee (" + item.currency_name + "):");
                                $('#panel_transaction_id_label').text("Transaction ID (" + item.currency_name + " or Internal):");

                                $('#panel_transaction_time').text(time);
                                $('#panel_transaction_currency').text(item.currency_name);
                                $('#panel_transaction_type').text(item.type);
                                $('#panel_transaction_amount').text(item.amount);
                                $('#panel_transaction_fee').text("N/A");
                                $('#panel_transaction_receiver_address').text(item.receiver_address);
                                $('#panel_transaction_id').text(item.txid);
                                $('#panel_transaction_status').text(item.status);

                                if (item.sender_address) {
                                    $('#panel_transaction_sender_address').text(item.sender_address);
                                } else {
                                    $('#panel_transaction_sender_address').text("N/A");
                                }

                                $('#panel_transaction_update_email').hide();

                                // Panel must be set to display "block" in order for animations to render
                                $("#panel_main_transaction").css({ display: "block" });
                                $("#panel_transaction").addClass("is-open");
                            });
                        }

                        payout_history = data.payout_transactions;
                        if (payout_history.length > 0) {
                            var trHTML = '<table class="ms-Table"><thead><tr><th>' + lang("Date") +
                                '</th><th>' + lang("Currency") + '</th><th>' + lang("Address") +
                                '</th><th>' + lang("Amount") + '</th><th>' + lang("Fee") +
                                '</th></tr></thead><tbody>';
                            $.each(payout_history, function (i, item) {
                                date = new Date(item.time).toLocaleDateString();
                                trHTML += '<tr><td>' + date + '</td><td>' + item.currency_name + '</td><td><div class="tooltip">' + item.receiver_address.substr(0, 6) +
                                    '...<span class="tooltiptext">' + item.receiver_address + '</span></td><td>' + item.amount + '</td><td>' + item.fee + '</td></tr>';
                            });
                            trHTML += '</tbody></table>';
                            $('#payout_history').html(trHTML);
                        }
                        $('#payout_history td').click(function () {
                            var id = $(this).parent().index();
                            var item = payout_history[id];
                            var time = new Date(item.time).toLocaleString();
                            var mailbox_item = Office.context.mailbox.item;

                            $('#panel_transaction_receiver_address_label').text("Receiver Address (" + item.currency_name + " or Email):");
                            $('#panel_transaction_amount_label').text("Amount (" + item.currency_name + "):");
                            $('#panel_transaction_fee_label').text("Fee (" + item.currency_name + "):");
                            $('#panel_transaction_id_label').text("Transaction ID (" + item.currency_name + " or Internal):");

                            $('#panel_transaction_time').text(time);
                            $('#panel_transaction_currency').text(item.currency_name);
                            $('#panel_transaction_type').text(item.type);
                            $('#panel_transaction_amount').text(item.amount);
                            $('#panel_transaction_fee').text(item.fee);
                            $('#panel_transaction_receiver_address').text(item.receiver_address);
                            $('#panel_transaction_id').text(item.txid);
                            $('#panel_transaction_status').text(item.status);

                            if (item.sender_address) {
                                $('#panel_transaction_sender_address').text(item.sender_address);
                            } else {
                                $('#panel_transaction_sender_address').text("N/A");
                            }

                            if (mailbox_item.body.getTypeAsync) {
                                $('#panel_transaction_update_email').show();
                            } else {
                                $('#panel_transaction_update_email').hide();
                            }

                            // Panel must be set to display "block" in order for animations to render
                            $("#panel_main_transaction").css({ display: "block" });
                            $("#panel_transaction").addClass("is-open");
                        });
                    }
                } else {
                    // error
                    currencies = [];

                    $('#pay_panel_messagebar').removeClass('ms-MessageBar--success');
                    $('#pay_panel_messagebar').addClass('ms-MessageBar--warning');
                    $('#pay_panel_messagebar_icon').removeClass('ms-Icon--receiptCheck');
                    $('#pay_panel_messagebar_icon').addClass('ms-Icon--refresh');

                    $('#info').text('Connecting...');
                    $('#deposit_address').text('');
                    $('#deposit_history').html("");
                    $('#payout_history').html("");

                    showNotification(lang(data.result));
                }
            },
            error: function (xhr, status, error) {
                // error
                currencies = [];

                $('#pay_panel_messagebar').removeClass('ms-MessageBar--success');
                $('#pay_panel_messagebar').addClass('ms-MessageBar--warning');
                $('#pay_panel_messagebar_icon').removeClass('ms-Icon--receiptCheck');
                $('#pay_panel_messagebar_icon').addClass('ms-Icon--refresh');

                $('#info').text('Connecting...');
                $('#deposit_address').text('');
                showNotification(lang("Failed to get user information"));
            }
        }).done(function (data) {
        });
    }

    function performUserSignUp(new_pwd, confirm_new_pwd, new_backup_email) {
        var new_md;
        var hashedNewPwd;
        var new_pwd_is_valid;
        var confirm_new_pwd_is_valid;
        var new_backup_email_is_valid;
        var both_new_pwd_and_email;

        new_pwd_is_valid = (new_pwd && new_pwd.length <= pwd_length);
        confirm_new_pwd_is_valid = (confirm_new_pwd && confirm_new_pwd.length <= pwd_length);
        new_backup_email_is_valid = validateEmail(new_backup_email);

        both_new_pwd_and_email = (new_pwd_is_valid && confirm_new_pwd_is_valid && new_pwd == confirm_new_pwd &&
            new_backup_email_is_valid);

        if (both_new_pwd_and_email) {
            new_md = new KJUR.crypto.MessageDigest({
                alg: "sha256",
                prov: "cryptojs"
            });

            if (new_md) {
                hashedNewPwd = new_md.digestString(new_pwd);

                Office.context.mailbox.getUserIdentityTokenAsync(performUserSignUpCb,
                    {
                        'hashed-new-pwd': hashedNewPwd,
                        'backup-email': new_backup_email,
                    });
            } else {
                $("#sign_up_msg").text(lang("Failed to create Finmail account. System error."));
            }
        } else {
            $("#sign_up_msg").text(lang("The input is invalid. Please try again."));
        }
    }

    function performUserSignUpCb(asyncResult) {
        var endpoint = appdomain + '/users';
        var token = asyncResult.value;
        var data = asyncResult.asyncContext;
        var profile = Office.context.mailbox.userProfile;
        var hashedNewPwd = data["hashed-new-pwd"];
        var backupEmail = data["backup-email"];

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'method': "sign-up",
                'token': token,
                'type': "msexch",
                'hashed-new-pwd': hashedNewPwd,
                'backup-email': backupEmail,
                'name': profile.displayName
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                if (data.result == "OK") {
                    $("#sign_up").prop('disabled', true);
                    $("#sign_up_msg").css({ color: "green" });
                    $("#sign_up_msg").text("Success! Please check backup mailbox for email verification." +
                        " You will be signed in automatically after 10 seconds.");
                    setTimeout(function () {
                        showPanel(true, false, false);
                        $("#sign_up_msg").css({ color: "red" });
                        getUserInfo();
                    }, 10000)
                } else {
                    $("#sign_up_msg").css({ color: "red" });
                    $("#sign_up_msg").text(lang(data.result));
                }
            },
            error: function (xhr, status, error) {
                // error
                $("#sign_up_msg").css({ color: "red" });
                $("sign_up_msg").text(lang("Failed to sign up due to network issue"));
            }
        }).done(function (data) {
        });
    }


    function performUserSignIn(sign_in_pwd) {
        var sign_in_md;
        var hashedSignInPwd;

        if (sign_in_pwd) {
            sign_in_md = new KJUR.crypto.MessageDigest({
                alg: "sha256",
                prov: "cryptojs"
            });

            if (sign_in_md) {
                hashedSignInPwd = sign_in_md.digestString(sign_in_pwd);
            } else {
                // TODO, If md can not be created, try the empty pwd
                hashedSignInPwd = "";
            }
        } else {
            hashedSignInPwd = "";
        }

        Office.context.mailbox.getUserIdentityTokenAsync(performUserSignInCb,
            {
                'hashed-sign-in-pwd': hashedSignInPwd
            });
    }

    function performUserSignInCb(asyncResult) {
        var endpoint = appdomain + '/auth';
        var token = asyncResult.value;
        var data = asyncResult.asyncContext;
        var profile = Office.context.mailbox.userProfile;
        var hashedSignInPwd = data["hashed-sign-in-pwd"];

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'method': "sign-in",
                'token': token,
                'type': "msexch",
                'hashed-sign-in-pwd': hashedSignInPwd
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                if (data.result == "OK") {
                    getUserInfo();

                    showPanel(true, false, false);
                } else {
                    $("#sign_in_pwd").text("");
                    $("#sign_in_pwd").focus();
                    if (hashedSignInPwd) {
                        $("#sign_in_msg").text(data.result);
                    }

                    showPanel(false, false, true);
                }
            },
            error: function (xhr, status, error) {
                //clearInterval(getUserInfo);

                $("#sign_in_pwd").text("");
                $("#sign_in_pwd").focus();
                $("#sign_in_msg").text("Failed to sign in due to network issue");

                showPanel(false, false, true);
            }
        }).done(function (data) {
        });
    }

    function performUserSignOut() {
        Office.context.mailbox.getUserIdentityTokenAsync(performUserSignOutCb);
    }

    function performUserSignOutCb(asyncResult) {
        var endpoint = appdomain + '/auth';
        var token = asyncResult.value;
        var profile = Office.context.mailbox.userProfile;

        $.ajax({
            url: endpoint,
            crossDomain: true,
            data: JSON.stringify({
                'email': profile.emailAddress,
                'method': "sign-out",
                'token': token,
                'type': "msexch"
            }),
            // headers: {'X-Requested-With': 'XMLHttpRequest'},
            contentType: 'application/json; charset=utf-8',
            type: 'POST',
            dataType: 'json',
            success: function (data) {
                if (data.result == "OK") {
                    // Panel must be set to display "block" in order for animations to render
                    $("#sign_in_pwd").text("");
                    $("#sign_in_pwd").focus();
                    $("#sign_in_msg").text("");

                    showPanel(false, false, true);
                } else {
                    showNotification(lang(data.result));
                }
            },
            error: function (xhr, status, error) {
                showNotification(lang("Failed to sign out due to network issue."));
            }
        }).done(function (data) {
        });
    }

    function getFee(currency) {
        var fee_list = [
            currency.fee_1h,
            currency.fee_2h,
            currency.fee_6h,
            currency.fee_12h
        ];

        var id = $('#fee option:selected').index();

        return fee_list[id];
    }

    function getCurrency() {
        var id = $('#currency option:selected').index();
        if (id < currencies.length) {
            return currencies[id];
        }
        return null;
    }

    // Ref: https://stackoverflow.com/questions/46155/how-to-validate-an-email-address-in-javascript?page=2&tab=votes#tab-top
    function validateEmail(email) {
        if (email) {
            var chrbeforAt = email.substr(0, email.indexOf('@'));
            if (!($.trim(email).length > 127)) {
                if (chrbeforAt.length >= 2) {
                    var re = /^(([^<>()[\]{}'^?\\.,!|//#%*-+=&;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?/;
                    //var re = /[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?/;
                    return re.test(email);
                } else {
                    return false;
                }
            } else {
                return false;
            }
        } else {
            return false;
        }
    }

    /**
     * With reference to jquery.fabric.js
   * Pivot Plugin
   *
   * Adds basic functionality to .ms-Pivot components.
   *
   * @param  {jQuery Object}  One or more .ms-Pivot components
   * @return {jQuery Object}  The same components (allows for chaining)
   */
    (function ($) {
        $.fn.initPivot = function () {

            /** Go through each pivot we've been given. */
            return this.each(function () {
                var $pivotContainer = $(this);

                /** When clicking/tapping a link, select it. */
                $pivotContainer.on('click', '.ms-Pivot-link', function (event) {
                    event.preventDefault();
                    var $this = $(this);
                    $this.siblings('.ms-Pivot-link').removeClass('is-selected');
                    $this.addClass('is-selected');

                    /* show/hide tab panel */
                    var $currpanel = $pivotContainer.nextAll('.ms-Pivot-panel-container').find('.ms-Pivot-panel.is-selected');
                    var $nextpanel = $('#' + $this.data('panel'));
                    $currpanel.css({
                        'z-index': 1
                    }).removeClass('is-selected');
                    $nextpanel.css({
                        'z-index': 100
                    }).addClass('is-selected');
                });
            });
        };
    })(jQuery);

    /**
     * Panel Plugin
     *
     * Adds basic demonstration functionality to .ms-Panel components.
     *
     * @param  {jQuery Object}  One or more .ms-Panel components
     * @return {jQuery Object}  The same components (allows for chaining)
     */
    (function ($) {
        $.fn.initPanel = function () {

            /** Go through each panel we've been given. */
            return this.each(function () {
                $panel = $(this);
                $panelMain = $panel.find(".ms-Panel-main");

                /** Hook to open the panel. */
                $(".ms-PanelAction-close").on("click", function () {

                    // Display Panel first, to allow animations
                    $panelMain.css({ display: "none" });
                    $panel.removeClass('is-open');

                    $panel.addClass("ms-Panel-animateOut");
                });

                $panelMain.on("animationend webkitAnimationEnd MSAnimationEnd", function (event) {
                    if (event.originalEvent.animationName === "fadeOut") {
                        // Hide and Prevent ms-Panel-main from being interactive
                        $panelMain.css({ display: "none" });
                        $panel.removeClass('is-open');
                    }
                });
            });
        };
    })(jQuery);

    /**
 * Dropdown Plugin
 *
 * Given .ms-Dropdown containers with generic <select> elements inside, this plugin hides the original
 * dropdown and creates a new "fake" dropdown that can more easily be styled across browsers.
 *
 * @param  {jQuery Object}  One or more .ms-Dropdown containers, each with a dropdown (.ms-Dropdown-select)
 * @return {jQuery Object}  The same containers (allows for chaining)
 */
    (function ($) {
        $.fn.initDropdown = function (command) {

            if (!!command) {
                switch (command.toLowerCase()) {
                    case "refresh":
                        console.log("refreshing");
                        var $dropdownWrapper = $(this),
                            $originalDropdown = $dropdownWrapper.children('.ms-Dropdown-select'),
                            $originalDropdownOptions = $originalDropdown.children('option'),
                            newDropdownTitle = "",
                            newDropdownItems = "";

                        $originalDropdownOptions.each(function (index, option) {

                            /** If the option is selected, it should be the new dropdown's title. */
                            if (option.selected) {
                                newDropdownTitle = option.text;
                            }

                            /** Add this option to the list of items. */
                            newDropdownItems += '<li class="ms-Dropdown-item' +
                                ((option.disabled) ? ' is-disabled"' : '') +
                                ((option.selected) ? ' is-selected"' : '"') +
                                '>' + option.text + '</li>';

                        });

                        // update title
                        $dropdownWrapper.find(".ms-Dropdown-title").html(newDropdownTitle);

                        // replace options content
                        var anchor = $dropdownWrapper.find(".ms-Dropdown-items");
                        anchor.children(".ms-Dropdown-item").remove();
                        anchor.append(newDropdownItems);
                        break;
                    default:
                        break;
                }

                return this;
            }

            /** Go through each dropdown we've been given. */
            return this.each(function () {

                var $dropdownWrapper = $(this),
                    $originalDropdown = $dropdownWrapper.children('.ms-Dropdown-select'),
                    $originalDropdownOptions = $originalDropdown.children('option'),
                    newDropdownTitle = '',
                    newDropdownItems = '',
                    newDropdownSource = '';

                /** Go through the options to fill up newDropdownTitle and newDropdownItems. */
                $originalDropdownOptions.each(function (index, option) {

                    /** If the option is selected, it should be the new dropdown's title. */
                    if (option.selected) {
                        newDropdownTitle = option.text;
                    }

                    /** Add this option to the list of items. */
                    newDropdownItems += '<li class="ms-Dropdown-item' +
                        ((option.disabled) ? ' is-disabled"' : '') +
                        ((option.selected) ? ' is-selected"' : '"') +
                        '>' + option.text + '</li>';

                });

                /** Insert the replacement dropdown. */
                newDropdownSource = '<span class="ms-Dropdown-title">' + newDropdownTitle + '</span><ul class="ms-Dropdown-items">' + newDropdownItems + '</ul>';
                $dropdownWrapper.append(newDropdownSource);

                function _openDropdown(evt) {
                    if (!$dropdownWrapper.hasClass('is-disabled')) {

                        /** First, let's close any open dropdowns on this page. */
                        $dropdownWrapper.find('.is-open').removeClass('is-open');

                        /** Stop the click event from propagating, which would just close the dropdown immediately. */
                        evt.stopPropagation();

                        /** Before opening, size the items list to match the dropdown. */
                        var dropdownWidth = $(this).parents(".ms-Dropdown").width();
                        $(this).next(".ms-Dropdown-items").css('width', dropdownWidth + 'px');

                        /** Go ahead and open that dropdown. */
                        $dropdownWrapper.toggleClass('is-open');
                        $('.ms-Dropdown').each(function () {
                            if ($(this)[0] !== $dropdownWrapper[0]) {
                                $(this).removeClass('is-open');
                            }
                        });

                        /** Temporarily bind an event to the document that will close this dropdown when clicking anywhere. */
                        $(document).bind("click.dropdown", function () {
                            $dropdownWrapper.removeClass('is-open');
                            $(document).unbind('click.dropdown');
                        });
                    }
                }

                /** Toggle open/closed state of the dropdown when clicking its title. */
                $dropdownWrapper.on('click', '.ms-Dropdown-title', function (event) {
                    _openDropdown(event);
                });

                /** Keyboard accessibility */
                $dropdownWrapper.on('keyup', function (event) {
                    var keyCode = event.keyCode || event.which;
                    // Open dropdown on enter or arrow up or arrow down and focus on first option
                    if (!$(this).hasClass('is-open')) {
                        if (keyCode === 13 || keyCode === 38 || keyCode === 40) {
                            _openDropdown(event);
                            if (!$(this).find('.ms-Dropdown-item').hasClass('is-selected')) {
                                $(this).find('.ms-Dropdown-item:first').addClass('is-selected');
                            }
                        }
                    }
                    else if ($(this).hasClass('is-open')) {
                        // Up arrow focuses previous option
                        if (keyCode === 38) {
                            if ($(this).find('.ms-Dropdown-item.is-selected').prev().siblings().length > 0) {
                                $(this).find('.ms-Dropdown-item.is-selected').removeClass('is-selected').prev().addClass('is-selected');
                            }
                        }
                        // Down arrow focuses next option
                        if (keyCode === 40) {
                            if ($(this).find('.ms-Dropdown-item.is-selected').next().siblings().length > 0) {
                                $(this).find('.ms-Dropdown-item.is-selected').removeClass('is-selected').next().addClass('is-selected');
                            }
                        }
                        // Enter to select item
                        if (keyCode === 13) {
                            if (!$dropdownWrapper.hasClass('is-disabled')) {

                                // Item text
                                var selectedItemText = $(this).find('.ms-Dropdown-item.is-selected').text();

                                $(this).find('.ms-Dropdown-title').html(selectedItemText);

                                /** Update the original dropdown. */
                                $originalDropdown.find("option").each(function (key, value) {
                                    if (value.text === selectedItemText) {
                                        $(this).prop('selected', true);
                                    } else {
                                        $(this).prop('selected', false);
                                    }
                                });
                                $originalDropdown.change();

                                $(this).removeClass('is-open');
                            }
                        }
                    }

                    // Close dropdown on esc
                    if (keyCode === 27) {
                        $(this).removeClass('is-open');
                    }
                });

                /** Select an option from the dropdown. */
                $dropdownWrapper.on('click', '.ms-Dropdown-item', function () {
                    if (!$dropdownWrapper.hasClass('is-disabled') && !$(this).hasClass('is-disabled')) {

                        /** Deselect all items and select this one. */
                        $(this).siblings('.ms-Dropdown-item').removeClass('is-selected');
                        $(this).addClass('is-selected');

                        /** Update the replacement dropdown's title. */
                        $(this).parents().siblings('.ms-Dropdown-title').html($(this).text());

                        /** Update the original dropdown. */
                        var selectedItemText = $(this).text();
                        $originalDropdown.find("option").each(function (key, value) {
                            if (value.text === selectedItemText) {
                                $(this).prop('selected', true);
                            } else {
                                $(this).prop('selected', false);
                            }
                        });
                        $originalDropdown.change();
                    }
                });

            });
        };
    })(jQuery);

    /**
 * List Item Plugin
 *
 * Adds basic demonstration functionality to .ms-ListItem components.
 *
 * @param  {jQuery Object}  One or more .ms-ListItem components
 * @return {jQuery Object}  The same components (allows for chaining)
 */
    (function ($) {
        $.fn.initListItem = function () {

            /** Go through each panel we've been given. */
            return this.each(function () {

                var $listItem = $(this);

                /** Detect clicks on selectable list items. */
                $listItem.on('click', '.js-toggleSelection', function () {
                    $(this).parents('.ms-ListItem').toggleClass('is-selected');
                });

            });

        };
    })(jQuery);
})();
Office.initialize = function () { };

function onNewMessageComposeHandler(event) {

    var signature = `
        <table style="width:320px; background-color:black" cellspacing="0" cellpadding="0">
            <tbody>
                <tr>
                    <td rowspan="3" style="border-left:3px solid blue; width:4px; background-color:rgb(236,255,188)">
                    </td>
                    <td valign="top" style="font-size:9pt; font-family:Arial; color:rgb(60,60,59); background-color:red">
                        <div style="background-color:rgb(236,255,188)"><b>Joanna Dark<br></b><em
                                style="font-size:9pt; color:rgb(60,60,59)">Software developer</em></div>
                    </td>
                </tr>
                <tr>
                    <td style="height:5px; background-color:rgb(236,255,188)"></td>
                </tr>
                <tr>
                    <td valign="top" style="font-size:9pt; font-family:Arial; background-color:red">
                        <div style="background-color:rgb(236,255,188)"><span style="font-size:12pt"><b
                                    style="color:rgb(189,39,45)">Sample Company</b><br></span>Lorem ipsum
                            dolor sit amet, consectetur adipiscing elit. Sed tempor semper enim at mollis. Maecenas faucibus
                            placerat pretium. Quisque in malesuada lorem. In ut tempor felis.</div>
                    </td>
                </tr>
            </tbody>
        </table>`;

    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function () { event.completed(); });
}


Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
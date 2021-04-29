// eslint-disable-next-line no-unused-vars
/* global window, console, setTimeout, document, Excel, Office, $ */

class Alerts {
    private delay = 2000;
    private alertContainer = '<div class="alert  position-fixed"></div>';
    private closeBtn = '<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>';

    error(message: string): void {
        const alert = this.build('Error', message);
        alert.addClass('alert-danger');
        alert.appendTo($('body'));
        setTimeout(() => { (<any>alert).alert('close') }, this.delay);
    }

    success(message: string): void {
        const alert = this.build('OK', message);
        alert.addClass('alert-success');
        alert.appendTo($('body'));
        setTimeout(() => { (<any>alert).alert('close') }, this.delay);
    }

    private build(type: string, message: string) {
        const alert = $(this.alertContainer);
        alert.append(this.closeBtn)
            .append('<strong>'+ type +': </strong>')
            .append('<span>'+ message +'</span>');
        return alert;
    }
}

export const alerts = new Alerts();
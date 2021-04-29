class Alerts {
    private delay = 2000;
    private alertContainer = `
    <div class="alert  position-fixed">
        <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
    </div>`;

    // eslint-disable-next-line no-undef
    error(message: string, delay?: number): JQuery<HTMLElement> {
        return this.build('Error', message, 'alert-danger', delay);
    }

    // eslint-disable-next-line no-undef
    success(message: string, delay?: number): JQuery<HTMLElement> {
        return this.build('OK', message, 'alert-success', delay);
    }

    // eslint-disable-next-line no-undef
    private build(type: string, message: string, alerClass: string, delay?: number): JQuery<HTMLElement> {
        const alert = $(this.alertContainer);
        alert.append('<strong>'+ type +': </strong>')
            .append('<span>'+ message +'</span>')
            .addClass(alerClass)
            .appendTo($('body'));
        setTimeout(() => { (<any>alert).alert('close') }, delay ?? this.delay);
        return alert;
    }
}

export const alerts = new Alerts();
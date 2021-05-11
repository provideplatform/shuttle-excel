class Alerts {
    private delay = 2000;
    private alertContainer = `
    <div class="alert  position-fixed">
        <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
    </div>`;

    // eslint-disable-next-line no-undef
    error(message: string | string[], delay?: number): JQuery<HTMLElement> {
        return this.build('Error', message, 'alert-danger', delay);
    }

    // eslint-disable-next-line no-undef
    success(message: string | string[], delay?: number): JQuery<HTMLElement> {
        return this.build('OK', message, 'alert-success', delay);
    }

    // eslint-disable-next-line no-undef
    warn(message: string | string[], delay?: number): JQuery<HTMLElement> {
        return this.build('WARN', message, 'alert-warning', delay);
    }

    // eslint-disable-next-line no-undef
    private build(type: string, message: string | string[], alerClass: string, delay?: number): JQuery<HTMLElement> {
        const alert = $(this.alertContainer);
        alert.append('<strong>'+ type +': </strong>');

        if (Array.isArray(message)) {
            message.forEach((val, index) => {
                const newline = index !== (message.length - 1) ? '<br/>' : '';
                alert.append('<span>'+ val +'</span>' + newline);
            });
        } else {
            alert.append('<span>'+ message +'</span>')
        }

        alert.addClass(alerClass)
            .appendTo($('body'));
        setTimeout(() => { (<any>alert).alert('close') }, delay ?? this.delay);
        return alert;
    }
}

export const alerts = new Alerts();

export function spinnerOff() {
    $('#overlay').hide();
}

export function spinnerOn() {
    $('#overlay').show();
}
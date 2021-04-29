import { AuthParams } from "./ident-client";

export class LoginFormData implements AuthParams {
    public email: string;
    public password: string;

    constructor($form) {
        this.email = <string>$form.find('#email').val();
        this.password = <string>$form.find('#password').val();
     }
    
    isValid(): boolean | string {
      if (this.email && this.password) {
        return true;
      } else {
        return "Email and password are required";
      }
    }
  }
import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import { Text, Button } from "@stardust-ui/react";
import * as microsoftTeams from "@microsoft/teams-js";
import "./signInPage.scss";

const SignInPage: React.FunctionComponent<RouteComponentProps> = props => {
    const errorMessage = "Please sign in to continue.";

    function onSignIn(): void {
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/signin-simple-start",
            successCallback: () => {
                console.log("Login succeeded!");
                window.location.href = "/dls";
            },
            failureCallback: (reason) => {
                console.log("Login failed: " + reason);
                window.location.href = "/errorpage";
            }
        });
    }

    return (
        <div className="sign-in-content-container">
            <Text
                content={errorMessage}
                size="medium"
            />
            <div className="space"></div>
            <Button content="Sign in" primary className="sign-in-button" onClick={onSignIn} />
        </div>
    );
};

export default SignInPage;

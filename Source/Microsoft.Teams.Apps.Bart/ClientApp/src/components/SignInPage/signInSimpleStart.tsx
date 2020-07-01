import React, { useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { getAuthenticationMetadata } from '../../apis/apiList';

const SignInSimpleStart: React.FunctionComponent = () => {
    useEffect(() => {
        microsoftTeams.initialize();
        microsoftTeams.getContext(context => {
            const windowLocationOriginDomain = window.location.origin.replace("https://", "");
            const loginHint = context.upn ? context.upn : "";
            getAuthenticationMetadata(windowLocationOriginDomain, loginHint).then(result => {
                window.location.assign(result.data);
            });
        });
    });

    return (
        <></>
    );
};

export default SignInSimpleStart;
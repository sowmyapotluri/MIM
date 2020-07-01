import * as React from 'react';
import { BrowserRouter, Route, Switch } from 'react-router-dom';
import './App.scss';
import { Provider, themes } from '@stardust-ui/react';
import * as microsoftTeams from "@microsoft/teams-js";
import { TeamsThemeContext, getContext, ThemeStyle } from 'msteams-ui-components-react';
import ErrorPage from "./components/ErrorPage/errorPage";
import SignInPage from "./components/SignInPage/signInPage";
import SignInSimpleStart from "./components/SignInPage/signInSimpleStart";
import SignInSimpleEnd from "./components/SignInPage/signInSimpleEnd";
import HelpPage from './components/HelpPage/HelpPage';
import SearchDetails from "./components/FindRoom/searchDetails"
import CreateIncident from "./components/IncidentCreation/CreateIncident"

export interface IAppState {
    theme: string;
    themeStyle: number;
}

class App extends React.Component<{}, IAppState> {

    constructor(props: {}) {
        super(props);
        this.state = {
            theme: "",
            themeStyle: ThemeStyle.Light,
        }
    }

    public componentDidMount = () => {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            let theme = context.theme || "";
            this.updateTheme(theme);
            this.setState({
                theme: theme
            });
        });

        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.updateTheme(theme);
            this.setState({
                theme: theme,
            }, () => {
                this.forceUpdate();
            });
        });
    }

    public setThemeComponent = () => {
        if (this.state.theme === "dark") {
            return (
                <Provider theme={themes.teamsDark}>
                    <div className="darkContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
        else if (this.state.theme === "contrast") {
            return (
                <Provider theme={themes.teamsHighContrast}>
                    <div className="highContrastContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        } else {
            return (
                <Provider theme={themes.teams}>
                    <div className="defaultContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
    }

    private updateTheme = (theme: string) => {
        if (theme === "dark") {
            this.setState({
                themeStyle: ThemeStyle.Dark
            });
        } else if (theme === "contrast") {
            this.setState({
                themeStyle: ThemeStyle.HighContrast
            });
        } else {
            this.setState({
                themeStyle: ThemeStyle.Light
            });
        }
    }

    public getAppDom = () => {
        const context = getContext({
            baseFontSize: 10,
            style: this.state.themeStyle
        });
        return (
            <TeamsThemeContext.Provider value={context}>
                <div className="appContainer">
                    <BrowserRouter>
                        <Switch>
                            <Route exact path="/errorpage" component={ErrorPage} />
                            <Route exact path="/errorpage/:id" component={ErrorPage} />
                            <Route exact path="/signin" component={SignInPage} />
                            <Route exact path="/signin-simple-start" component={SignInSimpleStart} />
                            <Route exact path="/signin-simple-end" component={SignInSimpleEnd} />
                            <Route exact path="/help" component={HelpPage} />
                            <Route exact path="/incident" component={CreateIncident} />
                        </Switch>
                    </BrowserRouter>
                </div>
            </TeamsThemeContext.Provider>
        );
    }

    public render(): JSX.Element {
        return (
            <div>
                {this.setThemeComponent()}
            </div>
        );
    }
}

export default App;
import * as React from "react";
import { override } from "@microsoft/decorators";
import * as _ from "lodash";

import styles from "./ProfileMeter.module.scss";
import IProfileMeterState from "./IProfileMeterState";
import IProfileMeterProps from "./IProfileMeterProps";
import IUserDetails from "./IUserDetails";
import { SPHttpClient, SPHttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";

import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { ProgressIndicator } from "office-ui-fabric-react/lib/ProgressIndicator";

export default class ProfileMeter extends React.Component<IProfileMeterProps, IProfileMeterState> {

    private _baseUrl: string;
    private _spHttpClient: SPHttpClient;

    constructor(props: IProfileMeterProps) {
        super(props);

        this.state = {
            score: 0,
            currentUser: null,
            showPanel: false
        };
    }

    public componentWillMount(): void {
        this._baseUrl = this.props.context.pageContext.web.absoluteUrl;
        this._spHttpClient = this.props.context.spHttpClient;
    }

    private async _getCurrentUserDetails(): Promise<IUserDetails> {

        const httpOptions: IHttpClientOptions = this._prepareHttpOptions();
        const userDetailsEndpoint: string =
            `${this._baseUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=AccountName,DisplayName,PictureUrl,UserProfileProperties`;

        return this._spHttpClient.get(userDetailsEndpoint, SPHttpClient.configurations.v1, httpOptions)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    private async _calculateScore(userDetails: IUserDetails): Promise<number> {

        let score: number = 0;

        var firstNameProperty = _.filter(userDetails.UserProfileProperties, { Key: "FirstName" })[0];
        var lastNameProperty = _.filter(userDetails.UserProfileProperties, { Key: "LastName" })[0];
        var aboutMeProperty = _.filter(userDetails.UserProfileProperties, { Key: "AboutMe" })[0];
        var designationProperty = _.filter(userDetails.UserProfileProperties, { Key: "SPS-JobTitle" })[0];
        var managerProperty = _.filter(userDetails.UserProfileProperties, { Key: "Manager" })[0];
        var birthdayProperty = _.filter(userDetails.UserProfileProperties, { Key: "SPS-Birthday" })[0];
        var schoolProperty = _.filter(userDetails.UserProfileProperties, { Key: "SPS-School" })[0];
        var interestsProperty = _.filter(userDetails.UserProfileProperties, { Key: "SPS-Interests" })[0];
        var skillsProperty = _.filter(userDetails.UserProfileProperties, { Key: "SPS-Skills" })[0];
        var pastProjectsProperty = _.filter(userDetails.UserProfileProperties, { Key: "SPS-PastProjects" })[0];

        var accountName = userDetails.AccountName ? userDetails.AccountName : null;
        var firstName = firstNameProperty["Value"] ? firstNameProperty["Value"] : null;
        var lastName = lastNameProperty["Value"] ? lastNameProperty["Value"] : null;
        var aboutMe = aboutMeProperty["Value"] ? aboutMeProperty["Value"] : null;
        var displayName = userDetails.DisplayName ? userDetails.DisplayName : null;
        var pictureUrl = userDetails.PictureUrl ? userDetails.PictureUrl : null;
        var designation = designationProperty["Value"] ? designationProperty["Value"] : null;
        var manager = managerProperty["Value"] ? managerProperty["Value"] : null;
        var birthday = birthdayProperty["Value"] ? birthdayProperty["Value"] : null;
        var school = schoolProperty["Value"] ? schoolProperty["Value"] : null;
        var interests = interestsProperty["Value"] ? interestsProperty["Value"] : [];
        var skills = skillsProperty["Value"] ? skillsProperty["Value"] : [];
        var pastProjects = pastProjectsProperty["Value"] ? pastProjectsProperty["Value"] : [];

        this.setState({
            currentUser: {
                AboutMe: aboutMe,
                DisplayName: displayName,
                PictureUrl: pictureUrl,
                Designation: designation,
                Manager: manager,
                Birthday: birthday,
                School: school,
                Interests: interests,
                Skills: skills,
                PastProjects: pastProjects
            }
        });

        for (var key in this.state.currentUser) {
            if(this.state.currentUser[key])
            {
                score += 10;
            }
        }

        return score;
    }

    private async _getProfileCompletenessScore(): Promise<number> {

        const userDetails: IUserDetails = await this._getCurrentUserDetails();

        if (userDetails) {
            const meterScore: number = await this._calculateScore(userDetails);
            return meterScore;
        }

        return -1;
    }

    public componentDidMount(): void {
        this._getProfileCompletenessScore().then(score => {
            this.setState({
                score: score
            });
        });
    }

    @override
    public render(): React.ReactElement<{}> {

        if (this.state.score === 0) {
            return <ProgressIndicator label="" description="" />;
        }

        return (
            <div className={styles.scoreHeader}>

                <button className={styles.dot} onClick={() => this.setState({ showPanel: true })}>
                    <span className={styles.score}>{`${this.state.score}%`}</span>
                </button>

                <Panel
                    isOpen={this.state.showPanel}
                    type={PanelType.smallFluid}
                    // tslint:disable-next-line:jsx-no-lambda
                    onDismiss={() => this.setState({ showPanel: false })}
                    headerText=""
                >
                    <dl>
                        <dt>
                            Profile Completeness Score - {this.state.score}%
                        </dt>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.DisplayName ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Display Name ({this.state.currentUser.DisplayName ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.PictureUrl ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Picture Url ({this.state.currentUser.PictureUrl ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.AboutMe ? "percentage-100" : "percentage-0"]}><span className={styles.text}>About Me ({this.state.currentUser.AboutMe ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.Birthday ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Birthday ({this.state.currentUser.Birthday ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.Designation ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Designation ({this.state.currentUser.Designation ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.School ? "percentage-100" : "percentage-0"]}><span className={styles.text}>School ({this.state.currentUser.School ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.Manager ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Manager ({this.state.currentUser.Manager ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.Interests ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Interests ({this.state.currentUser.Interests ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.Skills ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Skills ({this.state.currentUser.Skills ? "10%" : "0%"})</span></dd>
                        <dd className={styles["percentage"] + " " + styles[this.state.currentUser.PastProjects ? "percentage-100" : "percentage-0"]}><span className={styles.text}>Past Projects ({this.state.currentUser.PastProjects ? "10%" : "0%"})</span></dd>
                    </dl>

                </Panel >
            </div >
        );
    }

    private _prepareHttpOptions(): IHttpClientOptions {
        const httpOptions: IHttpClientOptions = {
            headers: this._prepareHeaders()
        };

        return httpOptions;
    }

    private _prepareHeaders(): Headers {
        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");

        return requestHeaders;
    }
}

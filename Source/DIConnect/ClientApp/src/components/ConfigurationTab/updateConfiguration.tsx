// <copyright file="updateConfiguration.tsx" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { TFunction } from 'i18next';
import { withTranslation, WithTranslation } from "react-i18next";
import { RouteComponentProps } from 'react-router-dom';
import { Flex, Text, Input, Alert, Button, Checkbox, Loader } from '@fluentui/react-northstar';
import { updateConfiguration, getERGConfiguration } from '../../apis/configurationSettingsApi';
import Constants from '../../constants/constants';
import './configurationTab.scss';

export interface IUpdateConfigurationState {
    loading: boolean,
    theme: string;
    ergDisplayButtonText: string;
    isERGDisplayButtonTextPresent: boolean;
    isFaqEnabled: boolean;
    isERGEnabledForTeam: boolean;
    submitLoading: boolean;
}

export interface UpdateConfigurationProps extends RouteComponentProps, WithTranslation {
}

class UpdateConfiguration extends React.Component<UpdateConfigurationProps, IUpdateConfigurationState> {
    readonly localize: TFunction;

    constructor(props: UpdateConfigurationProps) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            loading: true,
            theme:"",
            ergDisplayButtonText: this.localize('RegisterNewERGDefaultButtonText'),
            isERGDisplayButtonTextPresent: true,
            isFaqEnabled: false,
            isERGEnabledForTeam: false,
            submitLoading: false,
        }
    }

    public async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.setState({
                theme: context.theme!,
            });
        });
        let isERGEnabledForTeam = await this.getERGConfigurationDetails();
        this.setState({
            isERGEnabledForTeam: isERGEnabledForTeam,
            loading: false
        });
    }

    /**
   * Method to get ERG configuration details.
   */
    private getERGConfigurationDetails = async () => {
        try {
            const response = await getERGConfiguration();
            if (response.status === 200 && response.data) {
                this.setState({
                    ergDisplayButtonText: response.data.value
                });
                return response.data.isEnabled;
            }
        } catch (error) {
            // For first run experience we are limiting the create/request new group to global team. Handling 404 error to provide input from user.
            if (error.response.status === 404) {
                this.setState({
                    loading: false,
                });
            }
            else {
                throw error;
            }
        }
    }

    /**
    *Submit configuration details
    */
    private handleSubmit = async () => {
        if (!this.state.ergDisplayButtonText) {
            this.setState({ isERGDisplayButtonTextPresent: false });
            return;
        }

        this.setState({ submitLoading: true });
        let configurationData: object = {
            registerERGButtonDisplayText: this.state.ergDisplayButtonText,
            isERGCreationRestrictedToGlobalTeam: this.state.isERGEnabledForTeam
        };

        await updateConfiguration(configurationData);
        microsoftTeams.tasks.submitTask();
    }

    /**
     * Handling ERG for team enable check box change event.
     * @param isChecked | boolean value.
     */
    private onisERGEnabledForTeam = (isChecked: boolean): void => {
        this.setState({ isERGEnabledForTeam: !isChecked });
    }

    /**
    *Sets ERG display button text state.
    *@param value ERG display button text string
    */
    private onErgDisplayButtonTextChange = (value: string) => {
        this.setState({ ergDisplayButtonText: value, isERGDisplayButtonTextPresent: true });
    }

    /**
    *Returns text component containing error message for failed name field validation
    *@param {boolean} isValuePresent Indicates whether value is present
    */
    private getRequiredFieldError = (isValuePresent: boolean) => {
        if (!isValuePresent) {
            return (<Text content={this.localize('RequiredFieldMessage')} error size="small" />);
        }

        return (<></>);
    }

    public render(): JSX.Element {
        if (!this.state.loading) {
            return (
                <div className={this.state.theme === "default" ? "backgroundcolor" : ""} >
                    <Flex className="module-container" column>
                        <div className="configuration-seperation">
                            <Flex className="top-padding">
                                <Text className="margin-space" size="medium" weight="bold" content={this.localize('ERGConfigurationText')}/>
                            </Flex>
                            <Alert info className="top-padding"
                                content={<Flex className="top-padding"><Text className="margin-space" content={this.localize('ERGForTeamConfirmText')}/>
                                    <Flex.Item push>
                                        <Flex hAlign="end" >
                                            <Checkbox className="checkbox" toggle checked={this.state.isERGEnabledForTeam} onChange={() => this.onisERGEnabledForTeam(this.state.isERGEnabledForTeam)} />
                                        </Flex>
                                    </Flex.Item>
                                </Flex>}
                            />
                            <Flex className="top-padding">
                                <Text size="small" content={this.localize('ERGDiplayButtonTitleText')} className="margin-space" />
                            <Flex.Item push>
                                {this.getRequiredFieldError(this.state.isERGDisplayButtonTextPresent)}
                            </Flex.Item>
                            </Flex>
                            <Flex>
                                <Input
                                    className="between-space"
                                    maxLength={Constants.maxLengthERGButtonDisplayText}
                                    fluid
                                    value={this.state.ergDisplayButtonText}
                                    placeholder={this.localize('ERGDiplayButtonPlaceholderText')}
                                    onChange={(event: any) => this.onErgDisplayButtonTextChange(event.target.value)}
                                />
                            </Flex>
                        </div>
                        <Flex.Item push>
                            <Flex className="knowledge-base-footer" hAlign="end" >
                                <Button primary content={this.localize("SaveText")}
                                    onClick={this.handleSubmit} disabled={this.state.submitLoading}
                                    loading={this.state.submitLoading} />
                            </Flex>
                        </Flex.Item>
                    </Flex>
                </div>
            )
        }
        else {
            return <Loader />
        }
    }
}

const updateConfigurationWithTranslation = withTranslation()(UpdateConfiguration);
export default updateConfigurationWithTranslation;
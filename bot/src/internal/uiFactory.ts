import { MessageBuilder } from "@microsoft/teamsfx";
import { TabResponse, TabResponseCard } from "botbuilder";
import profileCard from "../adaptiveCards/profileCard.json";
import mailsCard from "../adaptiveCards/mailsCard.json";
import tasksCard from "../adaptiveCards/tasksCard.json";
import signOutCard from "../adaptiveCards/signOutCard.json";
import errorCard from "../adaptiveCards/errorCard.json";
import { ProfileCardData } from "../adaptiveCardModels/profileCardData";
import { MailsCardData } from "../adaptiveCardModels/mailsCardData";
import { TasksCardData } from "../adaptiveCardModels/tasksCardData";
import { Client } from '@microsoft/microsoft-graph-client';
import { Message, User, TodoTask } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import path from "path";

export class UiFactory {
    public async getSignInUI(signInLink: string): Promise<TabResponse> {
        return {
            tab: {
                type: "auth",
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: signInLink,
                            title: "Sign in to this app"
                        }
                    ]
                }
            }
        };
    }

    public async getSignOutUI(): Promise<TabResponse> {
        return {
            tab: {
              type: "continue",
              value: {
                  cards: [
                    {
                        card: MessageBuilder.attachAdaptiveCard(signOutCard, {}).attachments[0].content
                    }
                  ]
              },
            }
        };
    }

    public async getTabUi(tabEntityId: string, graphClient: Client): Promise<TabResponse> {
        const cards: TabResponseCard[] = [];
        if (tabEntityId === "dashboardTab") {
            cards.push(await this.getProfileCard(graphClient));
            cards.push(await this.getMailsCard(graphClient));
            cards.push(await this.getTasksCard(graphClient));
        }
        else if (tabEntityId === "outlookTab") {
            cards.push(await this.getMailsCard(graphClient));
            cards.push(await this.getTasksCard(graphClient));
        }
        else {
            cards.push(await this.getErrorCard());
        }
        return {
            tab: {
                type: "continue",
                value: {
                    cards: cards
                }
            }
        };
    }

    private async getErrorCard(): Promise<TabResponseCard> {
        return {
            card: MessageBuilder.attachAdaptiveCard(errorCard, {}).attachments[0].content
        };
    }

    private async getProfileCard(graphClient: Client): Promise<TabResponseCard> {
        let results = (await graphClient.api('/me').select('displayName,companyName,jobTitle,mobilePhone,mail').get()) as User;
       
        var imageString = '';
        let userImage = await graphClient.api('/me/photo/$value').get();
        if (userImage) {
            // Converting image of Blob type to base64 string for rendering as image.
            await userImage.arrayBuffer().then(result => {
                imageString = Buffer.from(result).toString('base64');
                if (imageString != '') {
                    // Writing file to Images folder to use as url in adaptive card
                    fs.writeFileSync(path.resolve(__dirname, './../images/profileimg.jpeg'), imageString, { encoding: 'base64' });
                }
            }).catch(error => { console.log(error) });
        }

        const profileCardData: ProfileCardData = {
            title: "Profile",
            displayName: results.displayName,
            profileImage: process.env.INITIATE_LOGIN_ENDPOINT.replace('/auth-start.html', '') + "/images/profileimg.jpeg",
            companyName: results.companyName,
            jobTitle: results.jobTitle,
            properties: [
                {
                    "key": "E-mail",
                    "value": results.mail
                },
                {
                    "key": "Phone",
                    "value": results.mobilePhone
                }
            ]
        };

        return {
            card: MessageBuilder.attachAdaptiveCard<ProfileCardData>(profileCard, profileCardData).attachments[0].content
        };
    }

    private async getMailsCard(graphClient: Client): Promise<TabResponseCard> {
        let results = await graphClient.api('/me/messages').select('subject,bodyPreview,sender,webLink').top(5).get();
        let mails:[Message] = results.value;
        
        const mailsCardData: MailsCardData = {
            "title": "Mails",
            "mails": mails
        };

        return {
            card: MessageBuilder.attachAdaptiveCard<MailsCardData>(mailsCard, mailsCardData).attachments[0].content
        };
    }

    private async getTasksCard(graphClient: Client): Promise<TabResponseCard> {
        let results = await graphClient.api('/me/todo/lists/Tasks/tasks').get();
        let tasks:[TodoTask] = results.value;

        const tasksCardData: TasksCardData = {
            "cardTitle": "Tasks",
            "tasks": tasks
        };

        return {
            card: MessageBuilder.attachAdaptiveCard<TasksCardData>(tasksCard, tasksCardData).attachments[0].content
        };
    }
}
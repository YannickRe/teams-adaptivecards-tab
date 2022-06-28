import { Activity, BotFrameworkAdapter, TabRequest, TabResponse, TabResponseCard, TabSubmit, TaskModuleRequest, TaskModuleResponse, TeamsActivityHandler, TurnContext } from "botbuilder";
import {
  CommandMessage,
  TeamsFxBotCommandHandler,
  TriggerPatterns,
  MessageBuilder,
} from "@microsoft/teamsfx";
import profileCard from "./adaptiveCards/profileCard.json";
import mailsCard from "./adaptiveCards/mailsCard.json";
import tasksCard from "./adaptiveCards/tasksCard.json";
import signOutCard from "./adaptiveCards/signOutCard.json";
import errorCard from "./adaptiveCards/errorCard.json";
import { ProfileCardData } from "./adaptiveCardModels/profileCardData";
import { MailsCardData } from "./adaptiveCardModels/mailsCardData";
import { TasksCardData } from "./adaptiveCardModels/tasksCardData";

/**
 * The `HelloWorldCommandHandler` registers a pattern with the `TeamsFxBotCommandHandler` and responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
export class AdaptiveCardsTabCommandHandler extends TeamsActivityHandler {

  protected async handleTeamsTabFetch(context: TurnContext, tabRequest: TabRequest): Promise<TabResponse> {
    // When the Bot Service Auth flow completes, context will contain a magic code used for verification.
    const magicCode =
    context.activity.value && context.activity.value.state
        ? context.activity.value.state
        : '';

    const bfAdapter = context.adapter as BotFrameworkAdapter;
    // Getting the tokenResponse for the user
    const tokenResponse = await bfAdapter.getUserToken(
        context,
        process.env.CONNECTION_NAME,
        magicCode
    );

    if (!tokenResponse || !tokenResponse.token) {
      // Token is not available, hence we need to send back the auth response
      const signInLink = await bfAdapter.getSignInLink(
          context,
          process.env.CONNECTION_NAME
      );
      // Retrieve the OAuth Sign in Link.


      // Generating and returning auth response.
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

    const profileCardData: ProfileCardData = {
      title: "Profile",
      displayName: "Yannick Reekmans",
      profileImage: "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
      companyName: "Qubix",
      jobTitle: "Office Development MVP",
      properties: [
          {
              "key": "E-mail",
              "value": "yannick@yrkmsn.onmicrosoft.com"
          },
          {
              "key": "Phone",
              "value": "+32 (472) 41 16 82"
          }
      ]
    };

    const mailsCardData: MailsCardData = {
      "title": "Mails",
      "mails": [
        {
            "subject": "Your weekly PIM digest for yrkmsn",
            "bodyPreview": "Here’s a summary of activities over the last seven days.Your weekly PIM digest for yrkmsnThanks for using Azure Active Directory Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:User a",
            "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAEMAABjzCQXAsbSR5fDh9706hiVAAJxTKf9AAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
            "sender": {
                "emailAddress": {
                    "name": "Microsoft Azure",
                    "address": "azure-noreply@microsoft.com"
                }
            }
        },
        {
            "subject": "Azure AD Identity Protection Weekly Digest",
            "bodyPreview": "See your Azure AD Identity Protection Weekly Digest reportAzure AD Identity Protection Weekly DigestyrkmsnNew risky users detected0New risky sign-ins detected(in real-time)0Privacy StatementMicrosoft Corporation",
            "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAEMAABjzCQXAsbSR5fDh9706hiVAAJxTKf9AAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
            "sender": {
                "emailAddress": {
                    "name": "Microsoft Azure",
                    "address": "azure-noreply@microsoft.com"
                }
            }
        },
        {
            "subject": "Your weekly PIM digest for yrkmsn",
            "bodyPreview": "Here’s a summary of activities over the last seven days.Your weekly PIM digest for yrkmsnThanks for using Azure Active Directory Privileged Identity Management (PIM). This weekly digest shows your PIM activities over the last seven days:User a",
            "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAEMAABjzCQXAsbSR5fDh9706hiVAAJxTKf9AAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
            "sender": {
                "emailAddress": {
                    "name": "Microsoft Azure",
                    "address": "azure-noreply@microsoft.com"
                }
            }
        },
        {
            "subject": "Last Reminder: Update your trusted root store by 30 June 2022 for IMDS Attested Data Users",
            "bodyPreview": "We'll begin updating Azure Instance Metadata Service TLS certificates in July 2022.Add DigiCert Global Root G2 to your Trusted Root Store if certificate pinning is in useYou're receiving this notice because you use Azure Instance Metadata Service ",
            "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAEMAABjzCQXAsbSR5fDh9706hiVAAJxTKf9AAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
            "sender": {
                "emailAddress": {
                    "name": "Microsoft Azure",
                    "address": "azure-noreply@microsoft.com"
                }
            }
        },
        {
            "subject": "Azure AD Identity Protection Weekly Digest",
            "bodyPreview": "See your Azure AD Identity Protection Weekly Digest reportAzure AD Identity Protection Weekly DigestyrkmsnNew risky users detected0New risky sign-ins detected(in real-time)0Privacy StatementMicrosoft Corporation",
            "webLink": "https://outlook.office365.com/owa/?ItemID=AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAEMAABjzCQXAsbSR5fDh9706hiVAAJxTKf9AAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
            "sender": {
                "emailAddress": {
                    "name": "Microsoft Azure",
                    "address": "azure-noreply@microsoft.com"
                }
            }
        }
      ]
    };

    const tasksCardData: TasksCardData = {
      "cardTitle": "Tasks",
      "tasks": [
          {
              "@odata.etag": "W/\"Y8wkFwLG0keXw4fe9OoYlQACcMDCNQ==\"",
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Task 2",
              "createdDateTime": "2022-06-27T23:25:22.4696762Z",
              "lastModifiedDateTime": "2022-06-27T23:25:22.5477937Z",
              "categories": [],
              "id": "AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAJxTaaFAAA=",
              "body": {
                  "content": "",
                  "contentType": "text"
              }
          },
          {
              "@odata.etag": "W/\"Y8wkFwLG0keXw4fe9OoYlQACcMDCLw==\"",
              "importance": "normal",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Task 1",
              "createdDateTime": "2022-06-27T23:25:19.5480112Z",
              "lastModifiedDateTime": "2022-06-27T23:25:19.6886204Z",
              "categories": [],
              "id": "AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAJxTaaEAAA=",
              "body": {
                  "content": "",
                  "contentType": "text"
              }
          },
          {
              "@odata.etag": "W/\"Y8wkFwLG0keXw4fe9OoYlQAAuvSY5w==\"",
              "importance": "low",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Task from Teams: Trigger flows from any message in Microsoft Teams",
              "createdDateTime": "2020-08-31T20:11:52.8226877Z",
              "lastModifiedDateTime": "2020-08-31T20:11:52.9282221Z",
              "categories": [],
              "id": "AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGbAAA=",
              "body": {
                  "content": "Microsoft released a new trigger in Power Automate, to trigger a flow on a specific message in Microsoft Teams. There are probably some very cool use cases we can create at clients. Read more on their blog: https://flow.microsoft.com/en-us/blog/trigger-flows-from-any-message-in-microsoft-teams/",
                  "contentType": "text"
              },
              "linkedResources@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('1b42366f-858f-417e-ad48-5dbfd7fe46c0')/todo/lists('Tasks')/tasks('AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGbAAA%3D')/linkedResources",
              "linkedResources": [
                  {
                      "webUrl": "https://teams.microsoft.com/l/message/19:c4ce8360f7c641509d078c96de6bf2ed@thread.skype/1598816878743",
                      "applicationName": "Microsoft Teams",
                      "displayName": "Trigger flows from any message in Microsoft Teams",
                      "externalId": "1598816878743",
                      "id": "616da95e-9521-4c6f-b2a8-86832605be6d"
                  }
              ]
          },
          {
              "@odata.etag": "W/\"Y8wkFwLG0keXw4fe9OoYlQAAuvSY4A==\"",
              "importance": "low",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Task from Teams: Trigger flows from any message in Microsoft Teams",
              "createdDateTime": "2020-08-31T19:56:50.6968979Z",
              "lastModifiedDateTime": "2020-08-31T19:56:50.7815257Z",
              "categories": [],
              "id": "AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGaAAA=",
              "body": {
                  "content": "Microsoft released a new trigger in Power Automate, to trigger a flow on a specific message in Microsoft Teams. There are probably some very cool use cases we can create at clients. Read more on their blog: https://flow.microsoft.com/en-us/blog/trigger-flows-from-any-message-in-microsoft-teams/",
                  "contentType": "text"
              },
              "linkedResources@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('1b42366f-858f-417e-ad48-5dbfd7fe46c0')/todo/lists('Tasks')/tasks('AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGaAAA%3D')/linkedResources",
              "linkedResources": [
                  {
                      "webUrl": "https://teams.microsoft.com/l/message/19:c4ce8360f7c641509d078c96de6bf2ed@thread.skype/1598816878743",
                      "applicationName": "Microsoft Teams",
                      "displayName": "Trigger flows from any message in Microsoft Teams",
                      "externalId": "1598816878743",
                      "id": "4fb2541f-6f00-482f-83fa-625337aff3cf"
                  }
              ]
          },
          {
              "@odata.etag": "W/\"Y8wkFwLG0keXw4fe9OoYlQAAuvSY2g==\"",
              "importance": "low",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Task from Teams: Trigger flows from any message in Microsoft Teams",
              "createdDateTime": "2020-08-31T19:55:34.6285367Z",
              "lastModifiedDateTime": "2020-08-31T19:55:34.8177074Z",
              "categories": [],
              "id": "AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGZAAA=",
              "body": {
                  "content": "Microsoft released a new trigger in Power Automate, to trigger a flow on a specific message in Microsoft Teams. There are probably some very cool use cases we can create at clients. Read more on their blog: https://flow.microsoft.com/en-us/blog/trigger-flows-from-any-message-in-microsoft-teams/",
                  "contentType": "text"
              },
              "linkedResources@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('1b42366f-858f-417e-ad48-5dbfd7fe46c0')/todo/lists('Tasks')/tasks('AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGZAAA%3D')/linkedResources",
              "linkedResources": [
                  {
                      "webUrl": "https://teams.microsoft.com/l/message/19:c4ce8360f7c641509d078c96de6bf2ed@thread.skype/1598816878743",
                      "applicationName": "Microsoft Teams",
                      "displayName": "Trigger flows from any message in Microsoft Teams",
                      "externalId": "1598816878743",
                      "id": "ead05306-1a85-44cd-8e9b-3fe8f7974165"
                  }
              ]
          },
          {
              "@odata.etag": "W/\"Y8wkFwLG0keXw4fe9OoYlQAAuvSDlg==\"",
              "importance": "low",
              "isReminderOn": false,
              "status": "notStarted",
              "title": "Task from Teams: Trigger flows from any message in Microsoft Teams",
              "createdDateTime": "2020-08-30T21:34:54.2831189Z",
              "lastModifiedDateTime": "2020-08-30T21:34:54.3548039Z",
              "categories": [],
              "id": "AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGYAAA=",
              "body": {
                  "content": "Microsoft released a new trigger in Power Automate, to trigger a flow on a specific message in Microsoft Teams. There are probably some very cool use cases we can create at clients. Read more on their blog: https://flow.microsoft.com/en-us/blog/trigger-flows-from-any-message-in-microsoft-teams/",
                  "contentType": "text"
              },
              "linkedResources@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('1b42366f-858f-417e-ad48-5dbfd7fe46c0')/todo/lists('Tasks')/tasks('AAMkAGNkM2Y1MzJlLThhNzctNDNkZS05ODBhLWY4NTg2NDdjNzczYwBGAAAAAAAJRbdiB9OZQoz5IT1HhDRtBwBjzCQXAsbSR5fDh9706hiVAAAAAAESAABjzCQXAsbSR5fDh9706hiVAAC7DqGYAAA%3D')/linkedResources",
              "linkedResources": [
                  {
                      "webUrl": "https://teams.microsoft.com/l/message/19:c4ce8360f7c641509d078c96de6bf2ed@thread.skype/1598816878743",
                      "applicationName": "Microsoft Teams",
                      "displayName": "Trigger flows from any message in Microsoft Teams",
                      "externalId": "1598816878743",
                      "id": "e7669388-d441-4047-a01a-f785f1b80b1d"
                  }
              ]
          }
      ]
    };

    const cards: TabResponseCard[] = [];
    if (tabRequest.tabContext.tabEntityId === "dashboardTab") {
      cards.push({
        card: MessageBuilder.attachAdaptiveCard<ProfileCardData>(profileCard, profileCardData).attachments[0].content
      });
      cards.push({
        card: MessageBuilder.attachAdaptiveCard<MailsCardData>(mailsCard, mailsCardData).attachments[0].content
      });
      cards.push({
        card: MessageBuilder.attachAdaptiveCard<TasksCardData>(tasksCard, tasksCardData).attachments[0].content
      });
    }
    else if (tabRequest.tabContext.tabEntityId === "outlookTab") {
      cards.push({
        card: MessageBuilder.attachAdaptiveCard<MailsCardData>(mailsCard, mailsCardData).attachments[0].content
      });
      cards.push({
        card: MessageBuilder.attachAdaptiveCard<TasksCardData>(tasksCard, tasksCardData).attachments[0].content
      });
    }
    else {
      cards.push({
        card: MessageBuilder.attachAdaptiveCard(errorCard, {}).attachments[0].content
      });
    }

    return {
      tab: {
        type: "continue",
        value: {
          cards: cards
        }
      }
    }
  }

  protected async handleTeamsTabSubmit(context: TurnContext, tabSubmit: TabSubmit): Promise<TabResponse> {
    const bfAdapter = context.adapter as BotFrameworkAdapter;
    if (tabSubmit.data.action === "signout") {
      await bfAdapter.signOutUser(context, process.env.ConnectionName);
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
  }
}

import {  BotFrameworkAdapter, CardFactory, TabRequest, TabResponse, TabResponseCard, TabSubmit, TaskModuleRequest, TaskModuleResponse, TeamsActivityHandler, TurnContext } from "botbuilder";
import {
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
import { UiFactory } from "./internal/uiFactory";
import { GraphClient } from "./internal/graphClient";

/**
 * The `HelloWorldCommandHandler` registers a pattern with the `TeamsFxBotCommandHandler` and responds
 * with an Adaptive Card if the user types the `triggerPatterns`.
 */
export class AdaptiveCardsTabCommandHandler extends TeamsActivityHandler {
  private uiFactory = new UiFactory();

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
      return await this.handleSignIn(bfAdapter, context);
    }

    let graphClient = new GraphClient(tokenResponse.token);
    return this.uiFactory.getTabUi(tabRequest.tabContext.tabEntityId, graphClient.Client);
  }

  private async handleSignIn(bfAdapter: BotFrameworkAdapter, context: TurnContext): Promise<TabResponse> {
    const signInLink = await bfAdapter.getSignInLink(
      context,
      process.env.CONNECTION_NAME
    );
    // Retrieve the OAuth Sign in Link.
    // Generating and returning auth response.
    return this.uiFactory.getSignInUI(signInLink);
  }

  protected async handleTeamsTabSubmit(context: TurnContext, tabSubmit: TabSubmit): Promise<TabResponse> {
    const bfAdapter = context.adapter as BotFrameworkAdapter;

    if (tabSubmit.data.action === "signout") {
      await bfAdapter.signOutUser(context, process.env.ConnectionName);
      return await this.uiFactory.getSignOutUI();
    }
    else {
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

      let graphClient = (new GraphClient(tokenResponse.token)).Client;

      if (tabSubmit.data.action === "marktaskcomplete") {
        await graphClient.api(`/me/todo/lists/Tasks/tasks/${tabSubmit.data.taskid}`).patch({
          "status": "completed"
        });
      }

      if (tabSubmit.data.action === "addtask") {
        await graphClient.api('/me/todo/lists/Tasks/tasks').post({
          "title": tabSubmit.data.addtask
        });
      }

      return this.uiFactory.getTabUi(tabSubmit.tabContext.tabEntityId, graphClient);
    }
  }
}

import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, MessageFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";

const { TaskModuleUIConstants } = require('../models/taskModuleUIConstants');
const { userdata } = require('../models/userdata');

const Actions = [
  TaskModuleUIConstants.AdaptiveCard,
  //TaskModuleUIConstants.CustomForm,
  //TaskModuleUIConstants.YouTube
];

export class ShowUserProfile extends SSOCommand {
  constructor() {
    super();
    this.matchPatterns = [/^\s*show\s*/];
    this.operationWithSSOToken = this.showUserInfo;
  }

  async showUserInfo(context: TurnContext, ssoToken: string) {
    await context.sendActivity("Retrieving user information from Microsoft Graph ...");

    // Call Microsoft Graph half of user
    const teamsfx = new TeamsFx().setSsoToken(ssoToken);
    const graphClient = createMicrosoftGraphClient(teamsfx, [
      "User.Read",
    ]);
    const me = await graphClient.api("/me").get();

    if (me) {
      userdata.input(me);
      // task invoke
      const reply = MessageFactory.list([
        test_class.getTaskModuleAdaptiveCardOptions()
      ]);
      await context.sendActivity(reply);

      // show user picture
      let photoBinary: ArrayBuffer;
      try {
        photoBinary = await graphClient
          .api("/me/photo/$value")
          .responseType(ResponseType.ARRAYBUFFER)
          .get();
      } catch {
        return;
      }

      const buffer = Buffer.from(photoBinary);
      const imageUri = "data:image/png;base64," + buffer.toString("base64");
      const card = CardFactory.thumbnailCard(
        "User Picture",
        CardFactory.images([imageUri])
      );
      await context.sendActivity({ attachments: [card] });

    } else {
      await context.sendActivity(
        "Could not retrieve profile information from Microsoft Graph."
      );
    }
  }
}

class Test {
  constructor() {}

getTaskModuleAdaptiveCardOptions() {
  const adaptiveCard = {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.0',
      type: 'AdaptiveCard',
      body: [
          {
              type: 'TextBlock',
              // text: 'Task Module Invocation from Adaptive Card',
              text: `SSO`,
              weight: 'bolder',
              size: 3
          }
      ],
      actions: Actions.map((cardType) => {
          return {
              type: 'Action.Submit',
              // title: cardType.buttonTitle,
              title: 'User Profile',
              data: { msteams: { type: 'task/fetch' }, data: cardType.id }
          };
      })
  };

  return CardFactory.adaptiveCard(adaptiveCard);
}

}

const test_class = new Test();



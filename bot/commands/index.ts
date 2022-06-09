import { BotCommand } from "../helpers/botCommand";
import { LearnCommand } from "./learn";
import { ShowUserProfile } from "./showUserProfile";
import { WelcomeCommand } from "./welcome";
import { MultiUserProfile } from "./multiUserProfile";

export const commands: BotCommand[] = [
  new LearnCommand(),
  new ShowUserProfile(),
  new WelcomeCommand(),
  new MultiUserProfile()
];

import IUWorkTracker from "./IUWorkTracker";
import { chatGPTSignOutPath, getChatGPTUser } from "./chatgpt-auth";

export default async function Home() {
  const chatGPTUser = await getChatGPTUser();
  // Computed here (server-side) rather than in the client bundle, so the account
  // popover never needs to import chatgpt-auth.ts's next/headers-dependent module.
  const chatGPTSignOutHref = chatGPTSignOutPath("/");
  return <IUWorkTracker chatGPTUser={chatGPTUser} chatGPTSignOutHref={chatGPTSignOutHref} />;
}

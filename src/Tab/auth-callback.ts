import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

broadcastResponseToMainFrame().catch(() => {
  const status = document.getElementById("auth-status");
  if (status) {
    status.textContent = "Sign-in could not be completed. Close this window and retry from the application.";
  }
});

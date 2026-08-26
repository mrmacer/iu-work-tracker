// @vitest-environment jsdom
import { cleanup, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, describe, expect, it, vi } from "vitest";
import IUWorkTracker from "../app/IUWorkTracker";
import { MemoryDataProvider } from "../lib/data-provider";

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
});

async function openQuickLog(provider = new MemoryDataProvider([])) {
  const user = userEvent.setup();
  render(<IUWorkTracker dataProvider={provider} />);
  const opener = (await screen.findAllByRole("button", { name: /^log work$/i }))[0];
  await user.click(opener);
  return { user, provider, opener };
}

async function reachOrbitStep(user: ReturnType<typeof userEvent.setup>) {
  await user.type(screen.getByRole("textbox", { name: /activity title/i }), "Development ORBIT interaction test");
  await user.selectOptions(screen.getByRole("combobox", { name: /activity type/i }), "Professional learning");
  await user.click(screen.getByRole("button", { name: /continue/i }));
  await user.click(screen.getByRole("button", { name: /continue/i }));
  await user.click(screen.getByRole("button", { name: /continue/i }));
  await user.click(screen.getByRole("button", { name: /continue/i }));
}

describe("Quick Log interaction and accessibility", () => {
  it("moves focus in, exposes step names, traps focus, and restores focus", async () => {
    const { user, opener } = await openQuickLog();
    const title = screen.getByRole("textbox", { name: /activity title/i });
    expect(document.activeElement).toBe(title);
    expect(screen.getByRole("button", { name: /step 1 of 5: what did you do.*, current step/i }).getAttribute("aria-current")).toBe("step");
    await user.tab({ shift: true });
    expect(document.activeElement).toBe(screen.getByRole("button", { name: /step 5 of 5/i }));
    await user.keyboard("{Escape}");
    expect(screen.queryByRole("dialog")).toBeNull();
    expect(document.activeElement).toBe(opener);
  });

  it("protects a dirty draft when Escape is pressed", async () => {
    const { user } = await openQuickLog();
    await user.type(screen.getByRole("textbox", { name: /activity title/i }), "Unsaved work");
    const confirm = vi.spyOn(window, "confirm").mockReturnValue(false);
    await user.keyboard("{Escape}");
    expect(screen.getByRole("dialog")).toBeTruthy();
    expect(confirm).toHaveBeenCalledOnce();
  });

  it("enables ORBIT by mouse and persists the reportable record", async () => {
    const { user, provider } = await openQuickLog();
    await reachOrbitStep(user);
    await user.click(screen.getByRole("checkbox", { name: /orbit reportable/i }));
    await user.click(screen.getByRole("button", { name: /primary deliverable B: PA STEELS/i }));
    await user.click(screen.getByRole("button", { name: /save & done/i }));
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeNull());
    const reread = await provider.getWorkRecords();
    expect(reread.status).toBe("success");
    if (reread.status === "success") expect(reread.value[0].orbit.reportable).toBe(true);
  });

  it("enables ORBIT and completes the final step by keyboard", async () => {
    const { user, provider } = await openQuickLog();
    await reachOrbitStep(user);
    const checkbox = screen.getByRole("checkbox", { name: /orbit reportable/i });
    checkbox.focus();
    await user.keyboard(" ");
    expect((checkbox as HTMLInputElement).checked).toBe(true);
    const deliverable = screen.getByRole("button", { name: /primary deliverable B: PA STEELS/i });
    deliverable.focus();
    await user.keyboard("{Enter}");
    const save = screen.getByRole("button", { name: /save & done/i });
    save.focus();
    await user.keyboard("{Enter}");
    await waitFor(() => expect(screen.queryByRole("dialog")).toBeNull());
    const reread = await provider.getWorkRecords();
    if (reread.status !== "success") throw new Error("Record reread failed");
    expect(reread.value[0].orbit.reportable).toBe(true);
  });
});

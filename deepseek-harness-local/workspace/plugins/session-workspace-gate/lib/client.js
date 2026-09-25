window.__ModuleLoader__.load({
  id: "dsh-session-workspace-gate",
  factory: (require) => {
    const module = { exports: {} };
    const React = require("react");
    const jsx = require("react/jsx-runtime");

    const sameOrChild = (base, candidate) => {
      const clean = (value) => String(value || "").replace(/[\\/]+/g, "/").replace(/\/$/, "").toLowerCase();
      const parent = clean(base);
      const child = clean(candidate);
      return Boolean(parent && child && (child === parent || child.startsWith(`${parent}/`)));
    };

    function GateDock({ session, workspaces, blocks, uiWorkspace }) {
      const [snapshot, setSnapshot] = React.useState(() => ({ sessions: workspaces.sessions(), workspaces: workspaces.list() }));
      const [status, setStatus] = React.useState("loading");
      const [retry, setRetry] = React.useState(0);
      const cwd = snapshot.sessions.byId[session.sessionId]?.cwd || "";
      const panel = { color: "var(--dsw-alias-label-primary, #f5f5f5)", background: "var(--dsw-alias-bg-module-platform, #242424)", border: "1px solid var(--dsw-alias-border-l2, #454545)", borderRadius: 12, padding: "14px 16px", marginBottom: 10, fontSize: 13, lineHeight: 1.5 };
      const step = { display: "flex", alignItems: "center", gap: 8, fontWeight: 600, fontSize: 14, marginBottom: 5 };
      const badge = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, borderRadius: 999, background: "#4c8dff", color: "#fff", fontSize: 12 };
      const doneBadge = { ...badge, background: "#278a62" };
      const button = { border: "1px solid #586174", borderRadius: 7, padding: "7px 12px", margin: "8px 8px 0 0", color: "#f5f7fb", background: "#303746", cursor: "pointer", fontSize: 13 };
      const primaryButton = { ...button, borderColor: "#4c8dff", background: "#2864c7" };

      React.useEffect(() => {
        const offSessions = workspaces.subscribeSessions(() => setSnapshot((old) => ({ ...old, sessions: workspaces.sessions() })));
        const offWorkspaces = workspaces.subscribeWorkspaces(() => setSnapshot((old) => ({ ...old, workspaces: workspaces.list() })));
        return () => { offSessions(); offWorkspaces(); };
      }, [workspaces]);

      React.useEffect(() => {
        let alive = true;
        if (!cwd) { setStatus("missing-cwd"); return undefined; }
        setStatus("checking");
        fetch("/api/onboarding?session_id=" + encodeURIComponent(session.sessionId), { credentials: "same-origin" })
          .then((response) => response.ok ? response.json() : Promise.reject(new Error(`状态读取失败（HTTP ${response.status}）`)))
          .then((record) => {
            if (!alive) return;
            const valid = record?.sessionDirectory === cwd && sameOrChild(cwd, record?.officeWorkspace);
            if (valid) { blocks.set(session.sessionId, undefined); setStatus("confirmed"); }
            else { blocks.set(session.sessionId, { reason: "请先完成会话目录选择，再确认办公工作区；完成前不能发送普通聊天。" }); setStatus("choose"); }
          })
          .catch((error) => {
            if (!alive) return;
            blocks.set(session.sessionId, { reason: "办公工作区确认状态读取失败，暂不能发送：" + error.message });
            setStatus("error");
          });
        return () => { alive = false; };
      }, [session.sessionId, cwd, blocks, retry]);

      if (status === "confirmed") return null;
      if (!cwd) return jsx.jsx("div", { role: "alert", style: panel, children: "当前会话没有可核实的会话目录，请点击新会话并先选择目录。" });
      if (status === "loading" || status === "checking") return jsx.jsx("div", { role: "status", style: panel, children: "正在读取办公工作区确认状态…" });
      if (status === "error") return jsx.jsxs("div", { role: "alert", children: [
        jsx.jsx("span", { children: "办公工作区确认状态不可用，暂不能发送消息。" }),
        jsx.jsx("button", { type: "button", onClick: () => setRetry((value) => value + 1), children: "重试读取状态" })
      ] });
      const confirm = (officeWorkspace) => {
        if (!officeWorkspace || !sameOrChild(cwd, officeWorkspace)) return;
        setStatus("saving");
        fetch("/api/onboarding", {
          method: "POST", credentials: "same-origin", headers: { "content-type": "application/json" },
          body: JSON.stringify({ session_id: session.sessionId, session_directory: cwd, office_workspace: officeWorkspace })
        }).then((response) => response.ok ? response.json() : response.json().catch(() => ({})).then((body) => Promise.reject(new Error(body.error || `确认失败（HTTP ${response.status}）`))))
          .then(() => { blocks.set(session.sessionId, undefined); setStatus("confirmed"); })
          .catch((error) => { blocks.set(session.sessionId, { reason: "办公工作区确认失败，暂不能发送：" + error.message }); setStatus("error"); });
      };
      const chooseChild = () => {
        setStatus("saving");
        Promise.resolve(uiWorkspace.pickDirectory()).then((path) => {
          if (!path) { setStatus("choose"); return null; }
          if (!sameOrChild(cwd, path) || path.toLowerCase() === cwd.toLowerCase()) throw new Error("请选择会话目录下的子目录");
          return workspaces.create({ path });
        }).then((workspace) => { if (workspace) confirm(workspace.path); }).catch((error) => { setStatus("choose"); blocks.set(session.sessionId, { reason: "请选择会话目录下的办公工作区：" + error.message }); });
      };
      return jsx.jsxs("div", { role: "group", "aria-label": "办公工作区确认", style: panel, children: [
        jsx.jsxs("div", { style: step, children: [jsx.jsx("span", { style: doneBadge, children: "✓" }), jsx.jsx("span", { children: "会话目录已选择" })] }),
        jsx.jsx("div", { style: { color: "#b9c1cf", paddingLeft: 30, marginBottom: 12, wordBreak: "break-all" }, children: cwd }),
        jsx.jsxs("div", { style: step, children: [jsx.jsx("span", { style: badge, children: "2" }), jsx.jsx("span", { children: "办公工作区使用刚选的会话目录吗？" })] }),
        jsx.jsx("button", { type: "button", style: primaryButton, disabled: status === "saving", onClick: () => confirm(cwd), children: "✓ 同意，使用同一目录" }),
        jsx.jsx("button", { type: "button", style: button, disabled: status === "saving", onClick: chooseChild, children: "⌁ 否，选会话目录下子目录" })
      ] });
    }

    function apply(ctx) {
      const sessions = ctx.get("sessions");
      const workspaces = ctx.get("workspaces");
      const blocks = ctx.get("conversation")?.blocks;
      if (!sessions || !workspaces || !blocks) throw new Error("session-workspace-gate: required client services unavailable");
      const uiWorkspace = ctx.get("uiWorkspace");
      if (!uiWorkspace) throw new Error("session-workspace-gate: uiWorkspace service unavailable");
      // ComposerBlocks stores one value per session. Preserve another plugin's
      // block while this onboarding gate is active, and restore it afterwards.
      const priorBlockSet = blocks.set;
      const rawBlockSet = priorBlockSet.bind(blocks);
      const outsideBlocks = new Map();
      const gateReasons = new Map();
      blocks.set = (sessionId, block) => {
        outsideBlocks.set(sessionId, block);
        rawBlockSet(sessionId, gateReasons.get(sessionId) || block);
      };
      const gateBlocks = { set(sessionId, block) {
        if (block) {
          if (!gateReasons.has(sessionId)) outsideBlocks.set(sessionId, blocks.storeFor(sessionId).getSnapshot());
          gateReasons.set(sessionId, block);
          rawBlockSet(sessionId, block);
        } else {
          gateReasons.delete(sessionId);
          rawBlockSet(sessionId, outsideBlocks.get(sessionId));
        }
      } };
      ctx.effect(() => () => {
        blocks.set = priorBlockSet;
        for (const sessionId of gateReasons.keys()) rawBlockSet(sessionId, outsideBlocks.get(sessionId));
      }, "session-workspace-gate: restore composer blocks");

      // ui-workspace's public startSession() deliberately reuses the current or
      // most recent Workspace. For this deployment, every New Session click
      // must first establish a fresh session directory. Keep the interception
      // at the public service boundary so the official native picker remains
      // the only directory chooser and the Host still validates the path.
      const originalStartSession = uiWorkspace.startSession.bind(uiWorkspace);
      let newSessionPending = false;
      const startFreshSession = () => {
        if (newSessionPending) return;
        newSessionPending = true;
        sessions.clear();
        Promise.resolve(uiWorkspace.pickDirectory()).then((path) => {
          if (!path) return;
          return workspaces.create({ path }).then((workspace) => sessions.create({ workspaceId: workspace.workspaceId })).then((sessionId) => uiWorkspace.openSession(sessionId));
        }).catch((error) => {
          // Picker cancellation is a normal no-op; actual errors remain visible
          // through the browser console and do not create a partially selected session.
          if (error) console.error("session-workspace-gate: directory selection failed", error);
        }).finally(() => { newSessionPending = false; });
      };
      uiWorkspace.startSession = startFreshSession;
      ctx.effect(() => () => { uiWorkspace.startSession = originalStartSession; }, "session-workspace-gate: restore new-session action");
      const data = {
        sessions: () => sessions.list.getSnapshot(),
        list: () => workspaces.list.getSnapshot(),
        subscribeSessions: (listener) => sessions.list.subscribe(listener),
        subscribeWorkspaces: (listener) => workspaces.list.subscribe(listener)
      };
      ctx.slots.inject("conversation.input.dock", () => ctx.slots.register({
        name: "conversation.input.dock", id: "session-workspace-gate", order: -100,
        inject: (sessionId) => {
          // Set the block synchronously when the session scope is mounted. The
          // React effect below only resolves/reconciles the persisted record.
          gateBlocks.set(sessionId, { reason: "请先完成会话目录选择，再确认办公工作区；完成前不能发送普通聊天。" });
          return {};
        }
      }, (props) => jsx.jsx(GateDock, { ...props, workspaces: data, blocks: gateBlocks, uiWorkspace: ctx.get("uiWorkspace") })));
      ctx.slots.inject("shell.overlay", () => ctx.slots.register({
        name: "shell.overlay", id: "session-workspace-gate-guide", order: 50
      }, () => {
        const current = React.useSyncExternalStore((listener) => sessions.list.subscribe(listener), () => sessions.list.getSnapshot().current);
        if (current !== undefined) return null;
        return jsx.jsxs("div", { role: "status", style: { position: "fixed", top: "58%", left: "50%", transform: "translate(-50%, -50%)", width: "min(88vw, 390px)", padding: "18px 20px", borderRadius: "12px", border: "1px solid var(--dsw-alias-border-l2, #454545)", background: "var(--dsw-alias-bg-module-platform, #242424)", color: "var(--dsw-alias-label-primary, #f5f5f5)", boxShadow: "0 8px 28px #0007", pointerEvents: "auto", zIndex: 20, textAlign: "left" }, children: [
          jsx.jsx("div", { style: { fontSize: 16, fontWeight: 650, marginBottom: 7 }, children: "开始新的工作" }),
          jsx.jsx("div", { style: { color: "#b9c1cf", fontSize: 13, lineHeight: 1.55, marginBottom: 13 }, children: "先选本地文件夹存放这次对话，随后确认办公工作区。" }),
          jsx.jsx("button", { type: "button", onClick: () => uiWorkspace.startSession(), style: { border: "1px solid #4c8dff", borderRadius: 7, padding: "8px 14px", color: "#fff", background: "#2864c7", cursor: "pointer", fontSize: 13 }, children: "选择会话目录" })
        ] });
      }));
    }
    module.exports.apply = apply;
    module.exports.inject = ["slots", "sessions", "workspaces", "conversation", "uiWorkspace"];
    return module.exports;
  }
});

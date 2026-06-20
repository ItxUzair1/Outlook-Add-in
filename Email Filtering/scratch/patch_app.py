import sys
c = open('src/taskpane/components/App.jsx', 'r', encoding='utf-8').read()

use_effect_find = '''    const mode = new URLSearchParams(window.location.search).get("mode");
    if (mode === "help" || mode === "search" || mode === "options" || mode === "onsend") {
      return;
    }'''

use_effect_replace = '''    const mode = new URLSearchParams(window.location.search).get("mode");
    if (mode === "file_multi") {
      try {
        const payloadStr = localStorage.getItem("multiEmailPayload");
        if (payloadStr) {
          const payload = JSON.parse(payloadStr);
          setMultiEmailItems(payload.items || []);
          setSubject(`Multiple Emails (${(payload.items || []).length})`);
        }
      } catch (e) {}
      return;
    }
    if (mode === "help" || mode === "search" || mode === "options" || mode === "onsend") {
      return;
    }'''

c = c.replace(use_effect_find, use_effect_replace)

file_email_find = '''    setIsFiled(false);
    setLoading(true);
    setMessage("Preparing to file...");
    abortControllerRef.current = new AbortController();'''

file_email_replace = '''    if (initialMode === "file_multi") {
      setIsFiled(false);
      setLoading(true);
      setMessage("Preparing to file multiple emails...");
      abortControllerRef.current = new AbortController();

      try {
        const selectedLocations = locations.filter((x) => selectedIds.includes(x.id));
        if (selectedLocations.length === 0) {
          throw new Error("Select at least one target location.");
        }
        
        const disconnected = selectedLocations.filter(loc => connectivityStatus[loc.id] === false);
        if (disconnected.length > 0) {
          const paths = disconnected.map(d => d.path.split("\\\\").pop()).join(", ");
          throw new Error(`Filing failed: Location(s) [${paths}] are disconnected. Please check your network connection.`);
        }

        let graphAccessToken = null;
        try {
          graphAccessToken = await getToken({ interactive: false });
        } catch (tokenErr) {
          console.warn("[App] No graph token available for multi-file:", tokenErr?.message);
        }

        if (koyoOptions.addFiledCategory !== false) {
          const categoryName = koyoOptions.filedCategoryName || "Filed by mailmanager (koyomail)";
          try {
            await ensureMasterCategory(categoryName, "Preset3");
          } catch (catErr) {
            console.warn("[App] Failed to ensure master category:", catErr.message);
          }
        }

        let filedCount = 0;
        let draftEmailCreatedOverall = false;
        let allSharingLinks = [];
        let accumulatedErrors = "";

        for (let i = 0; i < multiEmailItems.length; i++) {
          const item = multiEmailItems[i];
          setMessage(`Filing email ${i + 1} of ${multiEmailItems.length}...`);

          const validatedGraphAccessToken = (typeof graphAccessToken === "string" && graphAccessToken.length > 10) 
            ? graphAccessToken 
            : null;

          const payloadData = {
            itemId: item.itemId,
            subject: item.subject,
            graphAccessToken: validatedGraphAccessToken,
            isPartial: false,
            targetPaths: selectedLocations.map(l => l.folder || l.path),
            comment,
            attachmentsOption,
            markReviewed,
            sendLink,
            afterFiling: afterFiling || "none",
            addFiledCategory: koyoOptions.addFiledCategory !== false,
            filedCategoryName: koyoOptions.filedCategoryName || "Filed by mailmanager (koyomail)",
            useUtcTime: koyoOptions.useUtcTime || false,
            assistantCategories: koyoOptions.assistantCategories || "",
            duplicateStrategy: koyoOptions.duplicateStrategy || "rename",
            deleteEmptyFolders: koyoOptions.deleteEmptyFolders || false,
            filedFolderPrefix: koyoOptions.filedFolderPrefix || "*",
            applyReadOnly: koyoOptions.applyReadOnly || false
          };

          try {
            const response = await fileEmail(payloadData, { signal: abortControllerRef.current.signal });
            if (response.draftEmailCreated) draftEmailCreatedOverall = true;
            if (response.sharingLinks) allSharingLinks.push(...response.sharingLinks);
            if (response.postFilingError) accumulatedErrors += `[${item.subject}] ${response.postFilingError}\\n`;
            filedCount++;
            
            if (afterFiling && afterFiling !== "none") {
               if (Office.context.ui && Office.context.ui.messageParent) {
                 Office.context.ui.messageParent(JSON.stringify({ action: "afterFiling", value: afterFiling, itemId: item.itemId }));
               }
            }
          } catch (e) {
            console.error("Failed to file item", item.itemId, e);
            accumulatedErrors += `[${item.subject}] ${e.message}\\n`;
          }
        }
        
        let msg = `Successfully filed ${filedCount} of ${multiEmailItems.length} emails.`;
        if (accumulatedErrors) {
          msg += ` Some post-filing actions failed, check console.`;
          console.warn("Multi-file errors:", accumulatedErrors);
        }
        setMessage(msg);
        
        if (draftEmailCreatedOverall && allSharingLinks.length > 0) {
          openComposeWindow(allSharingLinks, "Multiple Emails");
        } else if (allSharingLinks.length > 0) {
          openComposeWindow(allSharingLinks, "Multiple Emails");
        }

        setIsFiled(true);
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        setMessage(`Filing failed: ${errorMsg}`);
      } finally {
        setLoading(false);
        abortControllerRef.current = null;
      }
      return;
    }

    setIsFiled(false);
    setLoading(true);
    setMessage("Preparing to file...");
    abortControllerRef.current = new AbortController();'''

c = c.replace(file_email_find, file_email_replace)

open('src/taskpane/components/App.jsx', 'w', encoding='utf-8').write(c)

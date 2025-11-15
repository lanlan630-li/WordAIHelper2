/* global Office, console */

// 全局变量，用来记录当前用户选择的功能和状态
let currentAction = '';
let isPreviewMode = false; // 【新增】状态开关，false=等待调用AI, true=等待写入文档

// 等待Office.js库加载完成
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // 为所有按钮绑定点击事件
        document.getElementById("polish-btn").onclick = () => showActionDialog('polish');
        document.getElementById("expand-btn").onclick = () => showActionDialog('expand');
        document.getElementById("other-btn").onclick = () => showActionDialog('other');
        document.getElementById("translate-btn").onclick = () => showActionDialog('translate');
        document.getElementById("explain-btn").onclick = () => showActionDialog('explain');

        // 为对话框的按钮绑定事件
        document.getElementById("confirm-btn").onclick = handleConfirm;
        document.getElementById("copy-btn").onclick = handleCopy;
        document.getElementById("cancel-btn").onclick = hideDialog;
    }
});

// 显示操作对话框
function showActionDialog(action) {
    currentAction = action;
    const dialog = document.getElementById("action-dialog");
    const title = document.getElementById("dialog-title");
    const input = document.getElementById("instruction-input");
    const preview = document.getElementById("preview-area");
    const confirmBtn = document.getElementById("confirm-btn");

    // 根据不同功能，设置不同的提示
    switch (action) {
        case 'polish': title.innerText = "润色文本"; input.placeholder = "请输入润色要求，例如：让语气更正式、更简洁"; break;
        case 'expand': title.innerText = "扩写文本"; input.placeholder = "请输入扩写要求，例如：增加更多细节、举一个例子"; break;
        case 'other': title.innerText = "自定义生成"; input.placeholder = "请输入您的任意需求，例如：写一封感谢信"; break;
        case 'translate': title.innerText = "翻译文本"; input.placeholder = "请输入目标语言，例如：翻译成英文"; break;
        case 'explain': title.innerText = "解释文本"; input.placeholder = "请输入解释要求，例如：用简单的语言解释"; break;
    }
    
    input.value = '';
    preview.innerText = '';
    dialog.style.display = 'block';

    // 【修正】每次打开对话框，都重置状态
    isPreviewMode = false;
    confirmBtn.innerText = "确认";
}

// 隐藏对话框
function hideDialog() {
    document.getElementById("action-dialog").style.display = 'none';
    // 【修正】每次关闭对话框，也重置状态
    isPreviewMode = false;
    document.getElementById("confirm-btn").innerText = "确认";
}

// 【修正后的核心函数】处理“确认”按钮的点击
async function handleConfirm() {
    const confirmBtn = document.getElementById("confirm-btn");
    const previewArea = document.getElementById("preview-area");

    if (!isPreviewMode) {
        // --- 步骤 1: 调用AI并预览 ---
        const instruction = document.getElementById("instruction-input").value;
        if (!instruction && currentAction !== 'other') {
            alert("请输入您的指令！");
            return;
        }

        confirmBtn.disabled = true; // 防止用户重复点击
        previewArea.innerText = "AI正在处理中，请稍候...";

        let selectedText = '';
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                context.load(selection, 'text');
                await context.sync();
                selectedText = selection.text;
            });
        } catch (e) {
            console.error("获取选中文本失败:", e);
        }

        let prompt = '';
        if (currentAction === 'other') {
            prompt = instruction;
        } else if (selectedText) {
            switch (currentAction) {
                case 'polish': prompt = `请帮我润色以下文本，要求：${instruction}。直接返回润色后的结果，不要有多余的解释。原文：\n${selectedText}`; break;
                case 'expand': prompt = `请帮我扩写以下文本，要求：${instruction}。直接返回扩写后的结果，不要有多余的解释。原文：\n${selectedText}`; break;
                case 'translate': prompt = `请帮我将以下文本翻译成${instruction}。直接返回翻译后的结果，不要有多余的解释。原文：\n${selectedText}`; break;
                case 'explain': prompt = `请帮我解释以下文本，要求：${instruction}。直接返回解释后的结果，不要有多余的解释。原文：\n${selectedText}`; break;
            }
        } else {
            previewArea.innerText = "请先选中一段文字再进行此操作！";
            confirmBtn.disabled = false;
            return;
        }

        const cleanedResult = await callAIAndClean(prompt);
        previewArea.innerText = cleanedResult;
        window.lastGeneratedContent = cleanedResult;

        // 切换到预览模式
        isPreviewMode = true;
        confirmBtn.innerText = "应用到文档";
        confirmBtn.disabled = false;

    } else {
        // --- 步骤 2: 将内容应用到文档 ---
        if (!window.lastGeneratedContent) {
            alert("没有可应用的内容！");
            return;
        }
        const contentToInsert = window.lastGeneratedContent;

        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                if (currentAction === 'other' || currentAction === 'polish' || currentAction === 'expand') {
                    selection.insertText(contentToInsert, "Replace");
                } else if (currentAction === 'translate' || currentAction === 'explain') {
                    selection.insertText(contentToInsert, "End");
                }
                await context.sync();
            });
            hideDialog(); // 操作完成后，隐藏对话框
        } catch (e) {
            console.error("写入Word失败:", e);
            alert("写入Word时出错，请重试。");
        }
    }
}

// 调用AI并清理结果
async function callAIAndClean(prompt) {
    // --- 在这里填写您的API Key ---
    const apiKey = "sk-dasbnnaq2xo7jnsd"; // <--- 把【】里的内容替换成您自己的Key！
    const apiUrl = "https://cloud.infini-ai.com/maas/v1/chat/completions";
    const modelName = "deepseek-v3.2-exp";

    const requestBody = { model: modelName, messages: [{ role: "user", content: prompt }] };

    try {
        const response = await fetch(apiUrl, {
            method: "POST", headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
            body: JSON.stringify(requestBody)
        });
        if (!response.ok) { const errorData = await response.json(); throw new Error(`请求失败: ${response.status} - ${errorData.error?.message || '未知错误'}`); }
        const data = await response.json();
        const aiReply = data.choices[0].message.content;
        return cleanUpText(aiReply);
    } catch (error) {
        console.error("调用AI时出错:", error);
        return `出错了: ${error.message}`;
    }
}

// 内容净化函数
function cleanUpText(rawText) {
    let cleaned = rawText;
    cleaned = cleaned.replace(/^#{1,6}\s+/gm, '').replace(/^\*\s+/gm, '').replace(/^\-\s+/gm, '').replace(/^\d+\.\s+/gm, '').replace(/^\>\s+/gm, '');
    cleaned = cleaned.replace(/\n\s*\n/g, '\n').trim();
    cleaned = cleaned.replace(/\n/g, '\n\n');
    return cleaned;
}

// 处理“复制”功能
function handleCopy() {
    if (!window.lastGeneratedContent) { alert("没有可复制的内容！"); return; }
    const textarea = document.createElement('textarea');
    textarea.value = window.lastGeneratedContent;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    alert("内容已复制到剪贴板！");
}

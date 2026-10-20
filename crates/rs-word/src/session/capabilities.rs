//! 能力探测：连接后按 `GetIDsOfNames` 与低风险读取判定 supported/untested。
//!
//! 只做只读探测；不执行未知宏、不改安全设置。

use crate::app::App;
use crate::types::{CapabilityInfo, CapabilityStatus};

fn cap(name: &str, ok: bool) -> CapabilityInfo {
    CapabilityInfo {
        name: name.to_string(),
        status: if ok {
            CapabilityStatus::Supported
        } else {
            CapabilityStatus::Untested
        },
        reason: if ok {
            None
        } else {
            Some("成员不存在或未验证".into())
        },
    }
}

/// 新增（V1/V2）能力：**成员存在不等于已验证**，一律 `untested`，真机实测后由
/// `CAPABILITY_MATRIX.md` 升级为 supported/partial/unsupported。
fn cap_untested(name: &str, exists: bool) -> CapabilityInfo {
    CapabilityInfo {
        name: name.to_string(),
        status: CapabilityStatus::Untested,
        reason: Some(if exists {
            "成员存在，行为未在真机验证".into()
        } else {
            "成员不存在".into()
        }),
    }
}

/// 探测当前会话能力。
pub fn probe(app: &App) -> Vec<CapabilityInfo> {
    let a = app.raw();
    let mut caps = vec![
        cap("connect", true),
        cap("application.documents", a.has_member("Documents")),
        cap("application.list_galleries", a.has_member("ListGalleries")),
        cap("vba.run", a.has_member("Run")),
        cap("history.undo_record", a.has_member("UndoRecord")),
        // 本地实现（无新 COM 依赖）的能力。
        cap("workflow.prepare_apply", true),
        cap("advanced.structure_diff", true),
    ];
    if let Ok(doc) = app.active_document() {
        let d = doc.raw();
        caps.extend([
            cap("document.range", d.has_member("Range")),
            cap("document.content", d.has_member("Content")),
            cap("document.find", d.has_member("Content")),
            cap("document.tables", d.has_member("Tables")),
            cap("document.sections", d.has_member("Sections")),
            cap("document.page_setup", d.has_member("PageSetup")),
            cap("document.export_pdf", d.has_member("ExportAsFixedFormat")),
            cap("document.track_revisions", d.has_member("TrackRevisions")),
            cap("document.revisions", d.has_member("Revisions")),
            cap("document.comments", d.has_member("Comments")),
            cap("document.vb_project", d.has_member("VBProject")),
            // V1/V2 新增项（成员存在 ≠ 已验证）。
            cap_untested("media.inline_shapes", d.has_member("InlineShapes")),
            cap_untested("media.shapes", d.has_member("Shapes")),
            cap_untested("advanced.footnotes", d.has_member("Footnotes")),
            cap_untested("advanced.endnotes", d.has_member("Endnotes")),
            cap_untested(
                "advanced.cross_reference",
                doc.content()
                    .map(|r| r.raw().has_member("InsertCrossReference"))
                    .unwrap_or(false),
            ),
            cap_untested(
                "table.format_v1",
                d.get_obj("Tables", "Tables")
                    .and_then(|t| t.get("Count"))
                    .is_ok(),
            ),
        ]);
        // WordArt 与自定义编号模板挂在别的对象上，单独取一次。
        if let Ok(shapes) = d.get_obj("Shapes", "Shapes") {
            caps.push(cap_untested(
                "media.wordart",
                shapes.has_member("AddTextEffect"),
            ));
        }
        if let Ok(project) = d.get_obj("VBProject", "VBProject") {
            caps.push(cap_untested(
                "vba.injection",
                project.has_member("VBComponents"),
            ));
        }
        if let Ok(templates) = d.get_obj("ListTemplates", "ListTemplates") {
            caps.push(cap_untested(
                "list.custom_template",
                templates.get("Count").is_ok(),
            ));
        }
    }
    caps
}

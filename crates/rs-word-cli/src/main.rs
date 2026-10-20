//! rs-word-cli —— WPS Writer / Microsoft Word COM 交互式 REPL 测试工具。

mod commands;
mod completion;

use std::cell::RefCell;
use std::rc::Rc;

use anyhow::Result;
use reedline::Signal;

use commands::State;
use completion::{editor, DynPrompt};

fn main() -> Result<()> {
    println!("rs-word-cli —— WPS Writer / Microsoft Word COM 交互测试工具");
    println!("直接 docs 可自动接管已打开的 WPS/Word 并列出文档；connect attach|new 建立连接。");
    println!("输入 help 查看全部命令。Tab 下方打开候选菜单，方向键导航，回车确认。");
    let state = Rc::new(RefCell::new(State::new()));
    let mut line_editor = editor(state.clone());
    loop {
        let prompt = DynPrompt {
            text: state.borrow().prompt(),
        };
        match line_editor.read_line(&prompt)? {
            Signal::Success(line) => {
                let line = line.trim();
                if line.is_empty() {
                    continue;
                }
                if !commands::execute(&state, line, &mut line_editor) {
                    break;
                }
            }
            Signal::CtrlC | Signal::CtrlD => break,
            _ => {}
        }
    }
    println!("已退出（连接的应用保持运行；关闭应用请用 killall 或手动关闭）。");
    Ok(())
}

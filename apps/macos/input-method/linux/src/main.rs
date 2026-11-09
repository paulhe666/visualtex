//! VisualTeX's Linux input method: an IBus engine (ibus-daemon starts it from its component,
//! visualtex.xml, with --ibus), with RIME pinyin and VisualTeX's formulas. `--xml` prints the
//! component for `ibus` to list.

mod engine;
mod ibus;
mod keys;
mod link;
mod log;
mod rime;
mod space;
mod x11;

use std::sync::Arc;
use tokio::sync::mpsc;
use zbus::zvariant::OwnedObjectPath;

const BUS_NAME: &str = "org.freedesktop.IBus.VisualTeX";

struct Factory {
    shared: Arc<engine::Shared>,
}

#[zbus::interface(name = "org.freedesktop.IBus.Factory")]
impl Factory {
    async fn create_engine(
        &self,
        #[zbus(object_server)] server: &zbus::ObjectServer,
        name: String,
    ) -> zbus::fdo::Result<OwnedObjectPath> {
        let path = self.shared.create();
        crate::log::write(&format!("engine {name} at {}", path.as_str()));
        server
            .at(path.clone(), engine::Engine { shared: self.shared.clone(), path: path.clone() })
            .await
            .map_err(|error| zbus::fdo::Error::Failed(error.to_string()))?;
        server
            .at(path.clone(), Service { shared: self.shared.clone(), path: path.clone() })
            .await
            .map_err(|error| zbus::fdo::Error::Failed(error.to_string()))?;
        Ok(path)
    }
}

/// ibus-daemon's Destroy for an engine it no longer needs.
struct Service {
    shared: Arc<engine::Shared>,
    path: OwnedObjectPath,
}

#[zbus::interface(name = "org.freedesktop.IBus.Service")]
impl Service {
    async fn destroy(&self, #[zbus(object_server)] server: &zbus::ObjectServer) {
        crate::log::write(&format!("engine destroyed {}", self.path.as_str()));
        self.shared.destroy(self.path.as_str());
        let _ = server.remove::<engine::Engine, _>(self.path.clone()).await;
        let _ = server.remove::<Service, _>(self.path.clone()).await;
    }
}

fn main() {
    let arguments: Vec<String> = std::env::args().collect();
    if arguments.iter().any(|argument| argument == "--version") {
        println!("{}", env!("CARGO_PKG_VERSION"));
        return;
    }
    let runtime = tokio::runtime::Builder::new_current_thread().enable_all().build().expect("a runtime");
    if let Err(error) = runtime.block_on(run()) {
        crate::log::write(&format!("stopped: {error}"));
        eprintln!("VisualTeX input method: {error}");
        std::process::exit(1);
    }
}

async fn run() -> Result<(), Box<dyn std::error::Error>> {
    let address = ibus::address().ok_or("IBus is not running (no IBus address)")?;
    let connection = zbus::connection::Builder::address(address.as_str())?.build().await?;
    crate::log::write(&format!("started {} pid={}", env!("CARGO_PKG_VERSION"), std::process::id()));
    // RIME deploys its tables in the background meanwhile
    let _ = rime::engine();
    let (incoming, mut messages) = mpsc::unbounded_channel();
    let link = link::Link::start(incoming);
    let shared = engine::Shared::new(connection.clone(), link);
    connection
        .object_server()
        .at("/org/freedesktop/IBus/Factory", Factory { shared: shared.clone() })
        .await?;
    connection.request_name(BUS_NAME).await?;
    // ibus-daemon's going away ends the engine (it starts another one when it comes back)
    let watched = connection.clone();
    let mut gone = Box::pin(async move {
        let Ok(proxy) = zbus::fdo::DBusProxy::new(&watched).await else { return };
        loop {
            tokio::time::sleep(std::time::Duration::from_secs(3)).await;
            // (ibus-daemon's bus answers NameHasOwner, not GetId)
            let ibus = zbus::names::BusName::try_from("org.freedesktop.IBus").expect("a bus name");
            if proxy.name_has_owner(ibus).await.is_err() {
                return;
            }
        }
    });
    loop {
        tokio::select! {
            message = messages.recv() => {
                let Some(message) = message else { break };
                engine::receive(shared.clone(), message).await;
            }
            _ = &mut gone => {
                crate::log::write("ibus-daemon gone");
                break;
            }
        }
    }
    Ok(())
}

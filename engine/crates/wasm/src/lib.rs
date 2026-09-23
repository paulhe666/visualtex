use wasm_bindgen::prelude::*;
use visualtex_engine_core::Runtime;

/// The JS object owns the runtime. Generated bindings copy request/response
/// strings; consumers must call free() when disposing the host. No raw pointers
/// or mutable document objects cross this boundary.
#[wasm_bindgen]
pub struct EngineRuntime {
    inner: Runtime,
}

#[wasm_bindgen]
impl EngineRuntime {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self { Self { inner: Runtime::default() } }

    pub fn dispatch(&mut self, request: &str) -> String {
        self.inner.dispatch_json(request)
    }

    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(&mut self, bytes: &[u8], face_index: u32) -> String {
        self.inner.register_font_json(bytes, face_index)
    }

    #[wasm_bindgen(getter, js_name = activeSessions)]
    pub fn active_sessions(&self) -> u32 { self.inner.active_sessions() as u32 }

    #[wasm_bindgen(getter, js_name = parkedSessions)]
    pub fn parked_sessions(&self) -> u32 { self.inner.parked_sessions() as u32 }

    #[wasm_bindgen(getter, js_name = parkedWeight)]
    pub fn parked_weight(&self) -> u32 { self.inner.parked_weight().min(u32::MAX as usize) as u32 }
}

impl Default for EngineRuntime {
    fn default() -> Self { Self::new() }
}

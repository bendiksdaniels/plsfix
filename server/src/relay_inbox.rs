//! `/api/inbox` routes: a sealed item dropped into a workspace, the items one
//! key can see, and the delete a deck makes once it has taken one. Owns the
//! JSON row shape and the link-id header; the rows live in `store_inbox` and
//! the write gates in `relay_gates`.
//! Invariant: a row is answered to the key that wrote it, and to no other.

use axum::{
    body::Bytes,
    extract::{Extension, Path, State},
    http::HeaderMap,
    response::IntoResponse,
    Json,
};
use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use serde::Serialize;

use crate::blocking::store_call;
use crate::relay::{is_link_id, now, ok_json, workspace_auth, Api, Refused, Reply, LINK_ID_HEADER};
use crate::relay_gates::{charge_bytes, full, ClientKey};
use crate::store_inbox::Posted;

#[derive(Serialize)]
struct InboxOut {
    id: String,
    #[serde(rename = "createdAt")]
    created_at: i64,
    blob: String,
}

pub(crate) async fn post_inbox(
    State(state): Api,
    Path(ws): Path<String>,
    client: Option<Extension<ClientKey>>,
    headers: HeaderMap,
    body: Bytes,
) -> Reply {
    let auth = workspace_auth(&headers, &ws)?;
    let id = headers
        .get(LINK_ID_HEADER)
        .and_then(|value| value.to_str().ok())
        .filter(|id| is_link_id(id))
        .ok_or(Refused::BadId)?
        .to_string();
    charge_bytes(&state, client.as_deref(), body.len())?;
    // No ownership check: the row is keyed by the writer's hash, so a foreign
    // key writes beside the pane's item rather than over it, and never sees it.
    let (space, link, item, at) = (ws.clone(), id.clone(), body.clone(), now());
    let posted = store_call(&state, "post_inbox", &id, move |store| {
        store.post_inbox(&space, &auth, &link, &item, at)
    })
    .await?;
    match posted {
        Posted::Stored => Ok(ok_json()),
        Posted::Full => Err(full("post_inbox", &id, &state)),
    }
}

pub(crate) async fn list_inbox(
    State(state): Api,
    Path(ws): Path<String>,
    headers: HeaderMap,
) -> Reply {
    let auth = workspace_auth(&headers, &ws)?;
    let (space, at) = (ws.clone(), now());
    let rows = store_call(&state, "list_inbox", &ws, move |store| {
        store.list_inbox(&space, &auth, at)
    })
    .await?;
    let out: Vec<InboxOut> = rows
        .into_iter()
        .map(|row| InboxOut {
            id: row.id,
            created_at: row.created_at,
            blob: URL_SAFE_NO_PAD.encode(row.blob),
        })
        .collect();
    Ok(Json(out).into_response())
}

pub(crate) async fn delete_inbox(
    State(state): Api,
    Path((ws, id)): Path<(String, String)>,
    headers: HeaderMap,
) -> Reply {
    let auth = workspace_auth(&headers, &ws)?;
    if !is_link_id(&id) {
        return Err(Refused::BadId);
    }
    let (space, link, at) = (ws.clone(), id.clone(), now());
    let deleted = store_call(&state, "delete_inbox", &id, move |store| {
        store.delete_inbox(&space, &auth, &link, at)
    })
    .await?;
    // A row this key did not write is invisible, so "not yours" is 404 here.
    if deleted {
        Ok(ok_json())
    } else {
        Err(Refused::Missing)
    }
}

//! Every refusal the relay routes can produce, and the one JSON body each one
//! answers with. Split out of `relay.rs` so the routes file stays under the
//! length limit; the routes decide which refusal, this decides how it reads.
//! Invariant: a refusal says the same thing on every route, and says nothing
//! about what the request carried.

use axum::{
    http::{header, StatusCode},
    response::{IntoResponse, Response},
    Json,
};

pub(crate) enum Refused {
    Unauthorized,
    BadId,
    BadRev,
    BadBody,
    TooManyItems,
    Forbidden,
    Missing,
    Full,
    RateLimited(u64),
    Store,
}

impl Refused {
    fn parts(&self) -> (StatusCode, &'static str) {
        match self {
            Refused::Unauthorized => (StatusCode::UNAUTHORIZED, "bearer required"),
            Refused::BadId => (StatusCode::BAD_REQUEST, "bad id"),
            Refused::BadRev => (StatusCode::BAD_REQUEST, "bad rev"),
            Refused::BadBody => (StatusCode::BAD_REQUEST, "bad body"),
            Refused::TooManyItems => (StatusCode::BAD_REQUEST, "too many items"),
            Refused::Forbidden => (StatusCode::FORBIDDEN, "another key owns this"),
            Refused::Missing => (StatusCode::NOT_FOUND, "not found"),
            Refused::Full => (StatusCode::INSUFFICIENT_STORAGE, "storage full"),
            Refused::RateLimited(_) => (StatusCode::TOO_MANY_REQUESTS, "too many requests"),
            Refused::Store => (StatusCode::INTERNAL_SERVER_ERROR, "store error"),
        }
    }
}

impl IntoResponse for Refused {
    fn into_response(self) -> Response {
        let (status, message) = self.parts();
        let body = Json(serde_json::json!({ "error": message }));
        match self {
            // The one refusal that can say when to come back: whole seconds,
            // the header a client is allowed to obey without parsing a body.
            Refused::RateLimited(seconds) => {
                (status, [(header::RETRY_AFTER, seconds.to_string())], body).into_response()
            }
            _ => (status, body).into_response(),
        }
    }
}

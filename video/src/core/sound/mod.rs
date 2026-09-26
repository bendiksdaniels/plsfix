//! sound: the soundtrack. cues = what the picture asks for; score = the music's notes (the groove
//! score, and the style switch); bed = the calm bed's notes;
//! instruments + sfx = the voices; bus = stereo buffers; mixdown = the master; track = a picked
//! song in place of the score.
pub mod bed;
pub mod bus;
pub mod cues;
pub mod instruments;
pub mod mixdown;
pub mod score;
pub mod sfx;
pub mod track;

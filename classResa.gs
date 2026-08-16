function createResaFromRecupMail(resa) {
    return  {
      idResa : resa[1] + " - " + resa[4],
      Statut : resa[2],
      Source : resa[3],
      Nom : resa[5],
      Email : "",
      Arrival : resa[6],
      Depart : resa[7],
      nbNuit : resa[8],
      nbRoom : 0,
      Room : "",
      Service : resa[10],
      Pax : resa[12],
      TotalMail : resa[18],
      Comment : "",
  };
}

function resaToTab (r) {
  return ([r.idResa,r.Statut,r.Source,r.Nom,r.Email,r.Arrival,r.Depart,r.nbNuit,r.nbRoom,
           r.Room,r.Service,r.Pax,r.TotalMail,r.Comment]);
}


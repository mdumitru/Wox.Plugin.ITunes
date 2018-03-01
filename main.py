# -*- coding: utf-8 -*-

import win32com.client as wc
from wox import Wox


class ITunes(Wox):

    def normalize_title(self, title):
        """Normalize track title.

        "ItemByName" is case-sensitive, so as a workaround we'll just
        capitalize the first letter in every word, for convenience (simply
        converting to title-case might not work for tracks with titles in
        all-caps - there it's the job of the user to provide the proper title).

        """
        return " ".join(w[0].capitalize() + w[1:] for w in title.split())

    def play_track(self, title):
        # FIXME error handling
        # TODO maybe attempt to simulate a case-insensitive search
        title = self.normalize_title(title)
        self.itunes = wc.Dispatch('iTunes.Application')
        song = self.itunes.LibrarySource.Playlists[0].Tracks.ItemByName(title)
        song.play()

    def query(self, query):
        results = []
        results.append({
            "Title": "ITunes",
            "SubTitle": "Query: {}".format(query),
            "IcoPath":"Images/itunes_logo.png",
            "ContextData": "ctxData",
            "JsonRPCAction": {
                "method": "play_track",
                "parameters": [query],
                "dontHideAfterAction": False
            }
        })
        return results


if __name__ == "__main__":
    ITunes()

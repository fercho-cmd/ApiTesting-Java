package com.HttpStatusCode;

import java.io.FileOutputStream;
import java.io.IOException;

import com.box.sdk.BoxAPIConnection;
import com.box.sdk.BoxFile;
import com.box.sdk.BoxFolder;
import com.box.sdk.BoxItem;

public class BoxService {

	public static void downloadBoxFile(String filename, String path) throws IOException {

		// String mySecretKey = System.getenv("API_KEY");
		// String folderID = System.getenv("FOLDER_ID");
		String mySecretKey = "";
		String folderID = "";
		BoxAPIConnection boxConnection = new BoxAPIConnection(mySecretKey); // API token
		BoxFolder folder = new BoxFolder(boxConnection, folderID); // Box folder ID
		String fileName = filename; //Name File from Box to be downloaded

		// For each file in folder
		// This is to avoid use the file ID because it can change
		for (BoxItem.Info itemInfo : folder) {
			if (itemInfo instanceof BoxFile.Info) {
				BoxFile.Info fileInfo = (BoxFile.Info) itemInfo;
				BoxFile file = new BoxFile(boxConnection, fileInfo.getID());
				String name = fileInfo.getName();
				if (name.toUpperCase().contains(fileName.toUpperCase())) {
					FileOutputStream stream = new FileOutputStream( path + "/" + name); // name of the file
					file.download(stream);
					stream.close();
					System.out.println("File downloaded");
					System.out.println(name);

				}
			} else if (itemInfo instanceof BoxFolder.Info) {
				BoxFolder.Info folderInfo = (BoxFolder.Info) itemInfo;
				folderInfo.getName();
			}
		}

	}

}

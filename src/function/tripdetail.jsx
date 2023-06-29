import axios from 'axios';

export const AddTripDetailFromExcel = async (data) => {
	return await axios.post("http://192.168.0.145:8080/api/posttripdetailfromexcel", data)
}
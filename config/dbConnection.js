import mongoose from 'mongoose'
import dotenv from 'dotenv'
dotenv.config()

const MONGODB_URI = 'mongodb+srv://f-ai:T5NUpnJvGuRpYO3R@f-ai.hae5tiq.mongodb.net/?retryWrites=true&w=majority&appName=f-ai'


const dbConnection = async () =>{
    try {
        const connectionInstance = await mongoose.connect(MONGODB_URI)
        console.log("dbconnected")
    } catch (e) {
        console.log(e.message)
        process.exit(1)
    }

}

export default dbConnection;
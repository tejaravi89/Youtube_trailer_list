 require 'httparty'
 require 'roo'

class Movie_details_bynames
    

  def initialize(movie,year) 
    apikey = "apikeygivenbyTheMoviedbOrg"   #https://www.themoviedb.org/
    uri_movie = CGI::escape(movie)  
    self.class.base_uri "https://api.themoviedb.org/3/search/movie?api_key=#{apikey}&query=#{uri_movie}&year=#{year}"
  end

	include HTTParty
  format :json  

  def get_info
	  self.class.get("")
	end	

end

class Get_trailer_byTMDB

  def initialize(tmdb_id) 
    apikey = "apikeygivenbyTheMoviedbOrg"   #https://www.themoviedb.org/
    self.class.base_uri "https://api.themoviedb.org/3/movie/#{tmdb_id}?api_key=#{apikey}&append_to_response=videos"
  end

  include HTTParty
  format :json  

  def get_info
    self.class.get("")
  end 

end

#both xls and xlsx formats can be used, movie names should be in first sheet and first column     
workbook = Roo::Spreadsheet.open('movies.xlsx')#Excel doc with movie names in the format - Toy Story(1995) 
workbook.default_sheet = workbook.sheets[0]
year_list = []
movies_names_list = workbook.column(1)
(0..movies_names_list.length-1).each do|i|
  year = movies_names_list[i].split("(",2)
    year[1] = year[1].chomp(')')
    year_list.push(year[1])
    movies_names_list[i] = year[0]
  end

      tmdb_list = []
      p "failed **************" if movies_names_list.length != year_list.length
      noOfMovies = movies_names_list.length-1
      (0..noOfMovies).each do |a|
        myobj = Movie_details_bynames.new(movies_names_list[a],year_list[a])
        t = Time.now
        sleep(t+0.25 - Time.now) if noOfMovies >= 40
        movie = myobj.get_info
        if movie != nil && movie["results"] != nil && movie["results"][0] != nil
          tmdb_list << movie["results"][0]["id"] 
        else
          tmdb_list << 0
        end
      end
      p tmdb_list

     
 
      result = []
      (0..tmdb_list.length-1).each do |a|
        if(tmdb_list[a] == 0) 
          onemovie = []
          onemovie << "Not Available"
          onemovie << "Not Available"
          onemovie << "Not Available"
          result << onemovie 
          next
        end

        myobj = Get_trailer_byTMDB.new(tmdb_list[a])
        movie = myobj.get_info
        (0..movie["videos"]["results"].length-1).each do |b|
        if movie != nil && movie["videos"] != nil && movie["videos"]["results"] != nil && movie["videos"]["results"][b]
         if movie["videos"]["results"][b]["type"] == "Trailer"
          onemovie = []
          onemovie << movie["title"]                        #movie name
          onemovie << movie["videos"]["results"][b]["key"]  # Youtube trailer key
          onemovie << movie["videos"]["results"][b]["name"] # name of the trailer
          result << onemovie  
          break
         end
        end
         if b == movie["videos"]["results"].length-1
          result << 0
        end
       end
      end

    (0..result.length-1).each do |b|
      if result[b][1] == "Not Available"
        err = "Check if there is typing error in movie name or year, may be trailer is not available"
        p err
        next
      end
      link = "https://youtu.be/"+result[b][1]
      puts link
    end




